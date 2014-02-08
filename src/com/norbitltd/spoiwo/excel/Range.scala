package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.util.{CellReference, CellRangeAddress}

object CellRange {

  val None = CellRange(0 -> 0, 0 -> 0)

}

case class CellRange(rowRange: (Int, Int), columnRange: (Int, Int)) {
  require(rowRange._1 <= rowRange._2, "First row can't be greater than the last row!")
  require(columnRange._1 <= columnRange._2, "First column can't be greate than the last column!")

  def convert(): org.apache.poi.ss.util.CellRangeAddress =
    new org.apache.poi.ss.util.CellRangeAddress(rowRange._1, rowRange._2, columnRange._1, columnRange._2)

}


object RowRange {

  val None = RowRange(0, 0)

}

case class RowRange(firstRowIndex: Int, lastRowIndex: Int) {
  require(firstRowIndex <= lastRowIndex, "First row index can't be greater than the last row index!")

  def convert(): CellRangeAddress = CellRangeAddress.valueOf("%d:%d".format(firstRowIndex, lastRowIndex))
}

object ColumnRange {

  val None = ColumnRange("A", "A")

  def apply(firstColumnIndex: Int, lastColumnIndex: Int) = {
    require(firstColumnIndex <= lastColumnIndex,
      "First column index can't be greater that the last column index!")
    ColumnRange(
      CellReference.convertNumToColString(firstColumnIndex),
      CellReference.convertNumToColString(lastColumnIndex)
    )
  }
}

case class ColumnRange(firstColumnName: String, lastColumnName: String) {
  def convert(): CellRangeAddress = CellRangeAddress.valueOf("%s:%s"
    .format(firstColumnName, lastColumnName))
}