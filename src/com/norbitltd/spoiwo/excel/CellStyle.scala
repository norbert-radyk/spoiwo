package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel._
import org.apache.poi.ss.usermodel.{VerticalAlignment, HorizontalAlignment, FillPatternType}

object CellStyle {

  val Default = CellStyle()

  val cache = collection.mutable.Map[XSSFWorkbook, collection.mutable.Map[CellStyle, XSSFCellStyle]]()
}

case class CellStyle(borders: CellBorder = CellBorder(),
                     dataFormat: CellDataFormat = CellDataFormat.Undefined,
                     font: Font = Font(),
                     fillPattern: FillPatternType = FillPatternType.NO_FILL,
                     fillForegroundColor: Color = Color.WHITE,
                     fillBackgroundColor: Color = Color.WHITE,
                     horizontalAlignment: HorizontalAlignment = HorizontalAlignment.GENERAL,
                     verticalAlignment: VerticalAlignment = VerticalAlignment.BOTTOM,
                     hidden: Boolean = false,
                     indention: Short = 0,
                     locked: Boolean = false,
                     rotation: Short = 0,
                     wrapText: Boolean = false) {

  def convert(cell: XSSFCell): XSSFCellStyle = convert(cell.getRow)

  def convert(row: XSSFRow): XSSFCellStyle = convert(row.getSheet)

  def convert(sheet: XSSFSheet): XSSFCellStyle = convert(sheet.getWorkbook)

  def convert(workbook: XSSFWorkbook): XSSFCellStyle = {
    val workbookCache = CellStyle.cache.getOrElseUpdate(workbook, collection.mutable.Map[CellStyle, XSSFCellStyle]())
    workbookCache.getOrElseUpdate(this, createCellStyle(workbook, this))
  }

  private def createCellStyle(workbook: XSSFWorkbook, cellStyle: CellStyle): XSSFCellStyle = {
    val cellStyle = workbook.createCellStyle()
    borders.applyTo(cellStyle)
    dataFormat.applyTo(workbook, cellStyle)
    cellStyle.setFont(font.convert(workbook))

    setFill(cellStyle)
    setAlignment(cellStyle)
    setProperties(cellStyle)
    cellStyle
  }

  private def setFill(cellStyle : XSSFCellStyle) {
    cellStyle.setFillPattern(fillPattern)
    cellStyle.setFillBackgroundColor(fillBackgroundColor.convert())
    cellStyle.setFillForegroundColor(fillForegroundColor.convert())
  }

  private def setAlignment(cellStyle : XSSFCellStyle) {
    cellStyle.setAlignment(horizontalAlignment)
    cellStyle.setVerticalAlignment(verticalAlignment)
  }

  private def setProperties(cellStyle : XSSFCellStyle) {
    cellStyle.setHidden(hidden)
    cellStyle.setIndention(indention)
    cellStyle.setLocked(locked)
    cellStyle.setRotation(rotation)
    cellStyle.setWrapText(wrapText)
  }

}
