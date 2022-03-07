package spoiwo.model.grid

import spoiwo.model.{Cell, Row}

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
      * So [Row(5), Row(), Row(1), Row(7), Row(6)] would possibly broke the output bcs 5 + 1 and 6 collides.
      * The given input should be all explicit or all implicit indexed.
      */
    def fixupIndexes: Seq[Row] = {
      RowHelper.fixupIndexes(rows)
    }

    def moveTo(x: Int, y: Int): Seq[Row] = {
      rows.fixupIndexes.map(r => r.withIndex(r.index.get + y).withCells(r.cells.map(c => c.withIndex(c.index.get + x))))
    }

    /**
      * The given rows list should be explicitly indexed
      */
    def merge: Seq[Row] = {
      rows.groupBy(_.index).map { case (_, v) =>
        val r = Row(v.flatMap(_.cells))
        v.headOption.flatMap(_.index).fold(r)(i => r.withIndex(i))
      }.toSeq.sortBy(_.index)
    }
  }
}
