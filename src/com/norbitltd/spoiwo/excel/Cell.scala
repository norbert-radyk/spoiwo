package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFRow}
import java.util.{Calendar, Date}
import org.apache.poi.ss.usermodel.FormulaError

object Cell extends Factory {

  //TODO Implementation for Hyperlink and Comment
  val defaultActive = false
  val defaultIndex = -1
  val defaultStyle = CellStyle.Default

  val Empty = apply("")

  def apply(value : String) : Cell = apply(value, defaultIndex, defaultStyle)
  def apply(value : String, index : Int) : Cell = apply(value, index, defaultStyle)
  def apply(value : String, style : CellStyle) : Cell = apply(value, defaultIndex, style)
  def apply(value : String, index : Int, style : CellStyle) : Cell = {

    val indexOption = wrap(index, defaultIndex)
    val styleOption = wrap(style, defaultStyle)

    if( value.startsWith("=")) {
      FormulaCell(value, indexOption, styleOption)
    } else if( value.contains("\n")) {
      StringCell(value, indexOption, Option(style.withWrapText(true)))
    } else {
      StringCell(value, indexOption, styleOption)
    }
  }

  def apply(value : Double) : Cell = apply(value, defaultIndex, defaultStyle)
  def apply(value : Double, index : Int) : Cell = apply(value, index, defaultStyle)
  def apply(value : Double, style : CellStyle) : Cell = apply(value, defaultIndex, style)
  def apply(value : Double, index : Int, style : CellStyle) : Cell =
    NumericCell(value, wrap(index, defaultIndex), wrap(style, defaultStyle))

  def apply(value : Int) : Cell = apply(value, defaultIndex, defaultStyle)
  def apply(value : Int, index : Int) : Cell = apply(value, index, defaultStyle)
  def apply(value : Int, style : CellStyle) : Cell = apply(value, defaultIndex, style)
  def apply(value : Int, index : Int, style : CellStyle) : Cell =
    apply(value.toDouble, index, style)

  def apply(value : Long) : Cell = apply(value, defaultIndex, defaultStyle)
  def apply(value : Long, index : Int) : Cell = apply(value, index, defaultStyle)
  def apply(value : Long, style : CellStyle) : Cell = apply(value, defaultIndex, style)
  def apply(value : Long, index : Int, style : CellStyle) : Cell =
    apply(value.toDouble, index, style)

  def apply(value : Boolean) : Cell = apply(value, defaultIndex, defaultStyle)
  def apply(value : Boolean, index : Int) : Cell = apply(value, index, defaultStyle)
  def apply(value : Boolean, style : CellStyle) : Cell = apply(value, defaultIndex, style)
  def apply(value : Boolean, index : Int, style : CellStyle) : Cell =
    BooleanCell(value, wrap(index, defaultIndex), wrap(style, defaultStyle))

  def apply(value : Date) : Cell = apply(value, defaultIndex, defaultStyle)
  def apply(value : Date, index : Int) : Cell = apply(value, index, defaultStyle)
  def apply(value : Date, style : CellStyle) : Cell = apply(value, defaultIndex, style)
  def apply(value : Date, index : Int, style : CellStyle) : Cell =
    DateCell(value, wrap(index, defaultIndex), wrap(style, defaultStyle))

  def apply(value : Calendar) : Cell = apply(value, defaultIndex, defaultStyle)
  def apply(value : Calendar, index : Int) : Cell = apply(value, index, defaultStyle)
  def apply(value : Calendar, style : CellStyle) : Cell = apply(value, defaultIndex, style)
  def apply(value: Calendar, index : Int, style : CellStyle) : Cell =
    CalendarCell(value, wrap(index, defaultIndex), wrap(style, defaultStyle))
}

sealed abstract class Cell(index : Option[Int], style : Option[CellStyle]) {

  def convert(row : XSSFRow) : XSSFCell

  private[excel] def convertInternal(row : XSSFRow)(initializeCell : XSSFCell => Unit) : XSSFCell = {
    val cellNumber = index.getOrElse(if( row.getLastCellNum < 0 ) 0 else row.getLastCellNum)
    val cell = row.createCell(cellNumber)
    style.foreach(s => cell.setCellStyle(s.convert(cell)))
    initializeCell(cell)
    cell
  }
}

case class StringCell(value : String, index : Option[Int], style : Option[CellStyle]) extends Cell(index, style) {
  override def convert(row : XSSFRow) = convertInternal(row) {
    cell : XSSFCell => cell.setCellValue(value)
  }
}

case class FormulaCell(formula: String, index : Option[Int], style : Option[CellStyle]) extends Cell(index, style) {
  override def convert(row : XSSFRow) = convertInternal(row) {
    cell : XSSFCell  => cell.setCellFormula(formula)
  }
}

case class ErrorValueCell(formulaError : FormulaError, index : Option[Int], style : Option[CellStyle]) extends Cell(index, style) {
  override def convert(row : XSSFRow) = convertInternal(row) {
    cell : XSSFCell  => cell.setCellErrorValue(formulaError)
  }
}

case class NumericCell(value : Double, index : Option[Int], style : Option[CellStyle]) extends Cell(index, style) {
  override def convert(row : XSSFRow) = convertInternal(row) {
    cell : XSSFCell  => cell.setCellValue(value)
  }
}

case class BooleanCell(value : Boolean, index : Option[Int], style : Option[CellStyle]) extends Cell(index, style) {
  override def convert(row : XSSFRow) = convertInternal(row) {
    cell : XSSFCell  => cell.setCellValue(value)
  }
}

case class DateCell(value : Date, index : Option[Int], style : Option[CellStyle]) extends Cell(index, style) {
  override def convert(row : XSSFRow) = convertInternal(row) {
    cell : XSSFCell  => cell.setCellValue(value)
  }
}

case class CalendarCell(value : Calendar, index : Option[Int], style : Option[CellStyle]) extends Cell(index, style) {
  override def convert(row: XSSFRow) = convertInternal(row) {
    cell : XSSFCell  => cell.setCellValue(value)
  }
}
