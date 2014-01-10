package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel._
import org.apache.poi.ss.usermodel.{VerticalAlignment, HorizontalAlignment, FillPatternType}

object CellStyle {

  val Default = CellStyle()

  val cache = collection.mutable.Map[XSSFWorkbook, collection.mutable.Map[CellStyle, XSSFCellStyle]]()
}

case class CellStyle(font : Font = Font(),
                       backgroundColor : Color = Color.WHITE,
                       borders : CellBorder = CellBorder(),
                       wrapText : Boolean = false,
                       dataFormat : CellDataFormat = CellDataFormat.Undefined,
                       horizontalAlignment : HorizontalAlignment = HorizontalAlignment.GENERAL,
                       verticalAlignment : VerticalAlignment = VerticalAlignment.BOTTOM) {

  def convert(cell : XSSFCell) : XSSFCellStyle = convert(cell.getRow)

  def convert(row : XSSFRow) : XSSFCellStyle = convert(row.getSheet)

  def convert(sheet : XSSFSheet) : XSSFCellStyle = convert(sheet.getWorkbook)

  def convert(workbook : XSSFWorkbook) : XSSFCellStyle = {
    val workbookCache = CellStyle.cache.getOrElseUpdate(workbook, collection.mutable.Map[CellStyle, XSSFCellStyle]())
    workbookCache.getOrElseUpdate(this, createCellStyle(workbook, this))
  }

  private def createCellStyle(workbook : XSSFWorkbook, cellStyle : CellStyle) : XSSFCellStyle = {
    val cellStyle = workbook.createCellStyle()
    cellStyle.setFont(font.convert(workbook))
    cellStyle.setFillBackgroundColor(backgroundColor.convert())
    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
    cellStyle.setWrapText(wrapText)
    cellStyle.setAlignment(horizontalAlignment)
    cellStyle.setVerticalAlignment(verticalAlignment)
    borders.convert(cellStyle)
    dataFormat.applyTo(workbook, cellStyle)
    cellStyle
  }

}
