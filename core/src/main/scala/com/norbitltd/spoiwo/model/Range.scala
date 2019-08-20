package com.norbitltd.spoiwo.model

//TODO conversion utility
import org.apache.poi.ss.util.CellReference

object CellRange {

  val None = CellRange(0 -> 0, 0 -> 0)

}

case class CellRange(rowRange: (Int, Int), columnRange: (Int, Int)) {
  require(rowRange._1 <= rowRange._2, "First row can't be greater than the last row!")
  require(columnRange._1 <= columnRange._2, "First column can't be greater than the last column!")
}

object RowRange {

  val None = RowRange(0, 0)

  def apply(rowRange: (Int, Int)): RowRange = RowRange(rowRange._1, rowRange._2)

}

case class RowRange(firstRowIndex: Int, lastRowIndex: Int) {
  require(firstRowIndex <= lastRowIndex, "First row index can't be greater than the last row index!")
}

object ColumnRange {

  val None = ColumnRange("A", "A")

  def apply(columnRange: (String, String)): ColumnRange = ColumnRange(columnRange._1, columnRange._2)

  def apply(firstColumnIndex: Int, lastColumnIndex: Int): ColumnRange = {
    require(firstColumnIndex <= lastColumnIndex, "First column index can't be greater that the last column index!")
    ColumnRange(
      CellReference.convertNumToColString(firstColumnIndex),
      CellReference.convertNumToColString(lastColumnIndex)
    )
  }
}

case class ColumnRange(firstColumnName: String, lastColumnName: String)
