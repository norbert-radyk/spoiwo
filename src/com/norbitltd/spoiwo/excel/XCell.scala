package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.usermodel.{Cell, CellStyle, Row}
import org.apache.poi.xssf.usermodel.XSSFRow

object XCell {

  val Empty = apply("", XCellStyle())

  def apply(value : String, cellStyle: XCellStyle) = if( value.startsWith("=")) {
    val formula = value.drop(1).trim
    FormulaCell(formula, cellStyle)
  } else {
    val valueBoundStyle = if( value.contains("\n"))
      cellStyle.copy(wrapText = true)
    else
      cellStyle
    StringCell(value, valueBoundStyle)
  }

  def apply(value : Double, cellStyle : XCellStyle) : XCell = NumericCell(value, cellStyle)

  def apply(value : Int, cellStyle : XCellStyle) : XCell = apply(value.toDouble, cellStyle)

  def apply(value : Long, cellStyle : XCellStyle) : XCell = apply(value.toDouble, cellStyle)




}

sealed trait XCell {

  def convert(row : XSSFRow) : Cell

  def convert(row : XSSFRow, style : XCellStyle)(initializeCell : Cell => Unit) : Cell = {
    val cellNumber = if( row.getLastCellNum < 0 ) 0 else row.getLastCellNum
    val cell = row.createCell(cellNumber)
    cell.setCellStyle(style.convert(cell))
    initializeCell(cell)
    cell
  }
}

case class StringCell(value : String, style : XCellStyle) extends XCell {
  override def convert(row : XSSFRow) = convert(row, style) {
    cell => cell.setCellValue(value)
  }
}

case class FormulaCell(formula: String, style : XCellStyle) extends XCell {
  override def convert(row : XSSFRow) = convert(row, style) {
    cell => cell.setCellFormula(formula)
  }
}

case class NumericCell(value : Double, style : XCellStyle) extends XCell {
  override def convert(row : XSSFRow) = convert(row, style) {
    cell => cell.setCellValue(value)
  }
}
