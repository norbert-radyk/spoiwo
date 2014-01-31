package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel._
import org.apache.poi.ss.usermodel.{VerticalAlignment, HorizontalAlignment, FillPatternType}

object CellStyle extends Factory {

  private lazy val defaultDataFormat = CellDataFormat(defaultPOICellStyle.getDataFormatString)
  private lazy val defaultFillPattern = defaultPOICellStyle.getFillPatternEnum
  private lazy val defaultFillForegroundColor = Color(defaultPOICellStyle.getFillForegroundXSSFColor)
  private lazy val defaultFillBackgroundColor = Color(defaultPOICellStyle.getFillBackgroundXSSFColor)
  private lazy val defaultHorizontalAlignment = defaultPOICellStyle.getAlignmentEnum
  private lazy val defaultVerticalAlignment = defaultPOICellStyle.getVerticalAlignmentEnum
  private lazy val defaultHidden = defaultPOICellStyle.getHidden
  private lazy val defaultIndention = defaultPOICellStyle.getIndention
  private lazy val defaultLocked = defaultPOICellStyle.getLocked
  private lazy val defaultRotation = defaultPOICellStyle.getRotation
  private lazy val defaultWrapText = defaultPOICellStyle.getWrapText

  private lazy val defaultBorders = CellBorder.Default
  private lazy val defaultFont = Font.Default

  val Default = CellStyle()

  def apply(borders: CellBorder = defaultBorders,
            dataFormat: CellDataFormat = defaultDataFormat,
            font: Font = defaultFont,
            fillPattern: FillPatternType = defaultFillPattern,
            fillForegroundColor: Color = defaultFillForegroundColor,
            fillBackgroundColor: Color = defaultFillBackgroundColor,
            horizontalAlignment: HorizontalAlignment = defaultHorizontalAlignment,
            verticalAlignment: VerticalAlignment = defaultVerticalAlignment,
            hidden: Boolean = defaultHidden,
            indention: Short = defaultIndention,
            locked: Boolean = defaultLocked,
            rotation: Short = defaultRotation,
            wrapText: Boolean = defaultWrapText): CellStyle =
    CellStyle(
      wrap(borders, defaultBorders),
      wrap(dataFormat, defaultDataFormat),
      wrap(font, defaultFont),
      wrap(fillPattern, fillPattern),
      wrap(fillForegroundColor, defaultFillForegroundColor),
      wrap(fillBackgroundColor, defaultFillBackgroundColor),
      wrap(horizontalAlignment, defaultHorizontalAlignment),
      wrap(verticalAlignment, defaultVerticalAlignment),
      wrap(hidden, defaultHidden),
      wrap(indention, defaultIndention),
      wrap(locked, defaultLocked),
      wrap(rotation, defaultRotation),
      wrap(wrapText, defaultWrapText)
    )

  private[excel] val cache = collection.mutable.Map[XSSFWorkbook, collection.mutable.Map[CellStyle, XSSFCellStyle]]()
}

case class CellStyle private[excel](
                                     borders: Option[CellBorder],
                                     dataFormat: Option[CellDataFormat],
                                     font: Option[Font],
                                     fillPattern: Option[FillPatternType],
                                     fillForegroundColor: Option[Color],
                                     fillBackgroundColor: Option[Color],
                                     horizontalAlignment: Option[HorizontalAlignment],
                                     verticalAlignment: Option[VerticalAlignment],
                                     hidden: Option[Boolean],
                                     indention: Option[Short],
                                     locked: Option[Boolean],
                                     rotation: Option[Short],
                                     wrapText: Option[Boolean]) {

  def withBorders(borders: CellBorder) =
    copy(borders = Option(borders))

  def withDataFormat(dataFormat: CellDataFormat) =
    copy(dataFormat = Option(dataFormat))

  def withFont(font: Font) =
    copy(font = Option(font))

  def withFillPattern(fillPattern: FillPatternType) =
    copy(fillPattern = Option(fillPattern))

  def withFillForegroundColor(fillForegroundColor: Color) =
    copy(fillForegroundColor = Option(fillForegroundColor))

  def withFillBackgroundColor(fillBackgroundColor: Color) =
    copy(fillBackgroundColor = Option(fillBackgroundColor))

  def withHorizontalAlignment(horizontalAlignment: HorizontalAlignment) =
    copy(horizontalAlignment = Option(horizontalAlignment))

  def withVerticalAlignment(verticalAlignment: VerticalAlignment) =
    copy(verticalAlignment = Option(verticalAlignment))

  def withHidden(hidden: Boolean) =
    copy(hidden = Option(hidden))

  def withIndention(indention: Short) =
    copy(indention = Option(indention))

  def withLocked(locked: Boolean) =
    copy(locked = Option(locked))

  def withRotation(rotation: Short) =
    copy(rotation = Option(rotation))

  def withWrapText(wrapText: Boolean) =
    copy(wrapText = Option(wrapText))

  def convert(cell: XSSFCell): XSSFCellStyle = convert(cell.getRow)

  def convert(row: XSSFRow): XSSFCellStyle = convert(row.getSheet)

  def convert(sheet: XSSFSheet): XSSFCellStyle = convert(sheet.getWorkbook)

  def convert(workbook: XSSFWorkbook): XSSFCellStyle = {
    val workbookCache = CellStyle.cache.getOrElseUpdate(workbook, collection.mutable.Map[CellStyle, XSSFCellStyle]())
    workbookCache.getOrElseUpdate(this, createCellStyle(workbook, this))
  }

  private def createCellStyle(workbook: XSSFWorkbook, cellStyle: CellStyle): XSSFCellStyle = {
    val cellStyle = workbook.createCellStyle()
    borders.foreach(b => b.applyTo(cellStyle))
    dataFormat.foreach(df => df.applyTo(workbook, cellStyle))
    font.foreach(f => f.convert(workbook))

    setFill(cellStyle)
    setAlignment(cellStyle)
    setProperties(cellStyle)
    cellStyle
  }

  private def setFill(cellStyle: XSSFCellStyle) {
    fillPattern.foreach(cellStyle.setFillPattern)
    fillBackgroundColor.foreach(c => cellStyle.setFillBackgroundColor(c.convert()))
    fillForegroundColor.foreach(c => cellStyle.setFillForegroundColor(c.convert()))
  }

  private def setAlignment(cellStyle: XSSFCellStyle) {
    horizontalAlignment.foreach(cellStyle.setAlignment)
    verticalAlignment.foreach(cellStyle.setVerticalAlignment)
  }

  private def setProperties(cellStyle: XSSFCellStyle) {
    hidden.foreach(cellStyle.setHidden)
    indention.foreach(cellStyle.setIndention)
    locked.foreach(cellStyle.setLocked)
    rotation.foreach(cellStyle.setRotation)
    wrapText.foreach(cellStyle.setWrapText)
  }

}
