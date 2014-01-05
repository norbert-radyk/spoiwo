package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFRow, XSSFSheet, XSSFWorkbook}

object XCellStyle {

  val Default = XCellStyle()

  val cache = collection.mutable.Map[Workbook, collection.mutable.Map[XCellStyle, CellStyle]]()
}

case class XCellStyle(font : XFont = XFont(),
                       backgroundColor : XColor = XColor.WHITE,
                       borderStyle : XCellBorder = XCellBorder(BorderStyle.NONE),
                       wrapText : Boolean = false) {

  def convert(cell : XSSFCell) : CellStyle = convert(cell.getRow)
  def convert(row : XSSFRow) : CellStyle = convert(row.getSheet)
  def convert(sheet : XSSFSheet) : CellStyle = convert(sheet.getWorkbook)
  def convert(workbook : XSSFWorkbook) : CellStyle = {
    val workbookCache = XCellStyle.cache.getOrElseUpdate(workbook, collection.mutable.Map[XCellStyle, CellStyle]())
    workbookCache.getOrElseUpdate(this, createCellStyle(workbook, this))
  }

  private def createCellStyle(workbook : XSSFWorkbook, cellStyle : XCellStyle) : CellStyle = {
    val cellStyle = workbook.createCellStyle()
    cellStyle.setFont(font.convert(workbook))
    cellStyle.setFillBackgroundColor(backgroundColor.convert())
    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
    cellStyle.setWrapText(wrapText)
    cellStyle.setBorderBottom(borderStyle.bottom)
    cellStyle.setBorderLeft(borderStyle.left)
    cellStyle.setBorderRight(borderStyle.right)
    cellStyle.setBorderTop(borderStyle.top)
    cellStyle
  }

}
