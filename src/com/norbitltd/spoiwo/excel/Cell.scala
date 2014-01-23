package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFRow}
import java.util.{Calendar, Date}
import org.apache.poi.ss.usermodel.FormulaError

object Cell {

  val Empty = apply("", CellStyle())

  def apply(value : String, cellStyle: CellStyle = CellStyle.Default) = if( value.startsWith("=")) {
    val formula = value.drop(1).trim
    FormulaCell(formula, cellStyle)
  } else {
    val valueBoundStyle = if( value.contains("\n"))
      cellStyle.copy(wrapText = true)
    else
      cellStyle
    StringCell(value, valueBoundStyle)
  }

  def apply(value : Double) : Cell = apply(value, CellStyle.Default)

  def apply(value : Double, cellStyle : CellStyle) : Cell = NumericCell(value, cellStyle)

  def apply(value : Int) : Cell = apply(value, CellStyle.Default)

  def apply(value : Int, cellStyle : CellStyle) : Cell = apply(value.toDouble, cellStyle)

  def apply(value : Long) : Cell = apply(value, CellStyle.Default)

  def apply(value : Long, cellStyle : CellStyle) : Cell = apply(value.toDouble, cellStyle)

  def apply(value : Boolean) : Cell = apply(value, CellStyle.Default)

  def apply(value : Boolean, cellStyle : CellStyle) : Cell = BooleanCell(value, cellStyle)

  def apply(value : Date) : Cell = apply(value, CellStyle.Default)

  def apply(value : Date, cellStyle : CellStyle) : Cell = DateCell(value, cellStyle)

  def apply(value: Calendar) : Cell = apply(value, CellStyle.Default)

  def apply(value: Calendar, cellStyle : CellStyle) : Cell = CalendarCell(value, cellStyle)
}

sealed trait Cell {

  def convert(row : XSSFRow) : XSSFCell

  def convert(row : XSSFRow, style : CellStyle)(initializeCell : XSSFCell => Unit) : XSSFCell = {
    val cellNumber = if( row.getLastCellNum < 0 ) 0 else row.getLastCellNum
    val cell = row.createCell(cellNumber)
    cell.setCellStyle(style.convert(cell))
    initializeCell(cell)
    cell
  }
}

case class StringCell(value : String, style : CellStyle) extends Cell {
  override def convert(row : XSSFRow) = convert(row, style) {
    cell => cell.setCellValue(value)
  }
}

case class FormulaCell(formula: String, style : CellStyle) extends Cell {
  override def convert(row : XSSFRow) = convert(row, style) {
    cell => cell.setCellFormula(formula)
  }
}

case class FormulaErrorCell(formulaError : FormulaError, style : CellStyle) extends Cell {
  override def convert(row : XSSFRow) = convert(row, style) {
    cell => cell.setCellErrorValue(formulaError)
  }
}

case class NumericCell(value : Double, style : CellStyle) extends Cell {
  override def convert(row : XSSFRow) = convert(row, style) {
    cell => cell.setCellValue(value)
  }
}

case class BooleanCell(value : Boolean, style : CellStyle) extends Cell {
  override def convert(row : XSSFRow) = convert(row, style) {
    cell => cell.setCellValue(value)
  }
}

case class DateCell(value : Date, style : CellStyle) extends Cell {
  override def convert(row : XSSFRow) = convert(row, style) {
    cell => cell.setCellValue(value)
  }
}

case class CalendarCell(value : Calendar, style : CellStyle) extends Cell {
  override def convert(row: XSSFRow) = convert(row, style) {
    cell => cell.setCellValue(value)
  }
}
