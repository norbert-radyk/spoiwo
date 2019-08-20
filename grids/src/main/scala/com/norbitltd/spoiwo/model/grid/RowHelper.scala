package com.norbitltd.spoiwo.model.grid

import com.norbitltd.spoiwo.model.{Cell, Row}

import scala.annotation.tailrec

object RowHelper {

  trait IndexFixupable[A] {
    def extractIndex(a: A): Option[Int]
    def fixIndex(a: A, newIdx: Int): A
    def fixOtherData(a: A): A
  }
  implicit val CellFixup: IndexFixupable[Cell] = new IndexFixupable[Cell] {
    override def extractIndex(a: Cell): Option[Int]   = a.index
    override def fixIndex(a: Cell, newIdx: Int): Cell = a.withIndex(newIdx)
    override def fixOtherData(a: Cell): Cell          = a
  }
  implicit val RowFixup: IndexFixupable[Row] = new IndexFixupable[Row] {
    override def extractIndex(a: Row): Option[Int]  = a.index
    override def fixIndex(a: Row, newIdx: Int): Row = a.copy(index = Some(newIdx))
    override def fixOtherData(a: Row): Row          = a.copy(cells = fixupIndexes(a.cells.toSeq))
  }

  def fixupIndexes[A](data: Seq[A])(implicit fixer: IndexFixupable[A]): Seq[A] = {
    data
      .foldLeft(List.empty[A] -> 0) {
        case ((list, idx), d) =>
          fixer
            .extractIndex(d)
            .fold {
              (fixer.fixOtherData(fixer.fixIndex(d, idx)) :: list) -> (idx + 1)
            } { i =>
              (fixer.fixOtherData(d) :: list) -> i
            }
      }._1.sortBy(fixer.extractIndex(_).getOrElse(0))
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
      rows.fixupIndexes.map(
        r => r.copy(index = Some(r.index.get + y), cells = r.cells.map(c => c.withIndex(c.index.get + x)))
      )
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
