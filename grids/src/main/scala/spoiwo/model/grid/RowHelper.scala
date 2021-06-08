package spoiwo.model.grid

import spoiwo.model.{Cell, Row}

import scala.annotation.tailrec

object RowHelper {

  def fixupIndexes(rows: Seq[Row]): Seq[Row] = {
    val indexedRows =
      if (rows.forall(_.index.isDefined)) {
        rows.sortBy(_.index.get)
      } else if (rows.forall(_.index.isEmpty)) {
        rows.zipWithIndex.map { case (rows, index) => rows.withIndex(index) }
      } else {
        throw new IllegalStateException("Not allowed to mix explicit and implicit indexing on cells!")
      }
    indexedRows.map(indexedRow => indexedRow.withCells(fixupCellIndexes(indexedRow.cells.toSeq)))
  }

  private def fixupCellIndexes(cells: Seq[Cell]): Seq[Cell] = {
    if (cells.forall(_.index.isDefined)) {
      cells.sortBy(_.index.get)
    } else if (cells.forall(_.index.isEmpty)) {
      cells.zipWithIndex.map { case (cell, index) => cell.withIndex(index) }
    } else {
      throw new IllegalStateException("Not allowed to mix explicit and implicit indexing on cells!")
    }
  }

  implicit class RowMergerHelper(rows: Seq[Row]) {

    /**
      * This function may not work on Rows with mixed explicit indexes if it has non-explicit indexes too!
      * So [Row(5), Row(), Row(1), Row(7), Row(6)] will possibly broke the output bcs 5 + 1 and 6 will collide!
      */
    def fixupIndexes: Seq[Row] = {
      RowHelper.fixupIndexes(rows)
    }

    def moveTo(x: Int, y: Int): Seq[Row] = {
      rows.fixupIndexes.map(r => r.withIndex(r.index.get + y).withCells(r.cells.map(c => c.withIndex(c.index.get + x))))
    }

    def merge: Seq[Row] = {
      def mergeRows(row1: Row, row2: Row) = {
        if (row1.index == row2.index) {
          (Row(index = row1.index.get, cells = row1.cells ++ row2.cells), true)
        } else {
          (row1, false)
        }
      }

      //This is probably slow as hell if you have huge number of rows
      @tailrec
      def merger(all: List[Row], acc: List[Row]): List[Row] = {
        all match {
          case h :: t =>
            val (newList, added) = acc.foldLeft((List.empty[Row], false)) {
              case ((acc, merged), row) =>
                val (newRow, mergedNow) = mergeRows(row, h)
                (newRow :: acc, mergedNow || merged)
            }
            val newestList = if (added) newList else h :: newList
            merger(t, newestList)
          case Nil => acc
        }
      }

      merger(rows.toList, List.empty)
    }
  }
}
