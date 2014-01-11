package com.norbitltd.spoiwo.excel

case class CellRangeAddress(rowRange : (Int, Int), columnRange : (Int, Int)) {
  require(rowRange._1 <= rowRange._2, "First row can't be greater than the last row!")
  require(columnRange._1 <= columnRange._2, "First column can't be greate than the last column!")

  def convert() : org.apache.poi.ss.util.CellRangeAddress =
    new org.apache.poi.ss.util.CellRangeAddress(rowRange._1, rowRange._2, columnRange._1, columnRange._2)

}
