package com.norbitltd.spoiwo.natures.xlsx

import org.apache.poi.ss.usermodel.{FillPatternType, HorizontalAlignment, VerticalAlignment, BorderStyle}
import com.norbitltd.spoiwo.model._
import org.apache.poi.xssf.usermodel._

object Model2XlsxConversions {

  private type Cache[K, V] = collection.mutable.Map[XSSFWorkbook, collection.mutable.Map[K, V]]

  private def Cache[K, V]() = collection.mutable.Map[XSSFWorkbook, collection.mutable.Map[K, V]]()

  private lazy val cellStyleCache = Cache[CellStyle, XSSFCellStyle]()
  private lazy val dataFormatCache = collection.mutable.Map[XSSFWorkbook, XSSFDataFormat]()

  implicit class XlsxBorderStyle(bs: CellBorderStyle) {
    def convertAsXlsx() = convertBorderStyle(bs)
  }

  implicit class XlsxColor(c: Color) {
    def convertAsXlsx() = convertColor(c)
  }

  implicit class XlsxCellFill(cf: CellFill) {
    def convertAsXlsx() = convertCellFill(cf)
  }

  implicit class XlsxCellStyle(cs: CellStyle) {
    def convertAsXlsx(cell: XSSFCell): XSSFCellStyle = convertAsXlsx(cell.getRow)
    def convertAsXlsx(row: XSSFRow): XSSFCellStyle = convertAsXlsx(row.getSheet)
    def convertAsXlsx(sheet: XSSFSheet): XSSFCellStyle = convertAsXlsx(sheet.getWorkbook)
    def convertAsXlsx(workbook: XSSFWorkbook) : XSSFCellStyle = convertCellStyle(cs, workbook)
  }

  implicit class XlsxHorizontalAlignment(ha: CellHorizontalAlignment) {
    def convertAsXlsx() = convertHorizontalAlignment(ha)
  }

  implicit class XlsxVerticalAlignment(va: CellVerticalAlignment) {
    def convertAsXlsx() = convertVerticalAlignment(va)
  }

  private def convertBorderStyle(borderStyle: CellBorderStyle): BorderStyle = {
    import CellBorderStyle._
    import BorderStyle._

    borderStyle match {
      case DashDot => DASH_DOT
      case DashDotDot => DASH_DOT_DOT
      case Dashed => DASHED
      case Dotted => DOTTED
      case Double => DOUBLE
      case Hair => HAIR
      case Medium => MEDIUM
      case MediumDashDot => MEDIUM_DASH_DOT
      case MediumDashDotDot => MEDIUM_DASH_DOT_DOTC
      case None => NONE
      case SlantedDashDot => SLANTED_DASH_DOT
      case Thick => THICK
      case Thin => THIN
      case CellBorderStyle(id) =>
        throw new Exception("Unsupported option for XLSX conversion with id=%d".format(id))
    }
  }

  private def convertColor(color: Color): XSSFColor = new XSSFColor(
    Array[Byte](color.r.toByte, color.g.toByte, color.b.toByte)
  )

  private def convertCellBorders(borders: CellBorders, style: XSSFCellStyle) {
    borders.leftStyle.foreach(s => style.setBorderLeft(convertBorderStyle(s)))
    borders.leftColor.foreach(c => style.setLeftBorderColor(convertColor(c)))
    borders.bottomStyle.foreach(s => style.setBorderBottom(convertBorderStyle(s)))
    borders.bottomColor.foreach(c => style.setBottomBorderColor(convertColor(c)))
    borders.rightStyle.foreach(s => style.setBorderRight(convertBorderStyle(s)))
    borders.rightColor.foreach(c => style.setRightBorderColor(convertColor(c)))
    borders.topStyle.foreach(s => style.setBorderTop(convertBorderStyle(s)))
    borders.topColor.foreach(c => style.setTopBorderColor(convertColor(c)))
  }

  private def convertCellDataFormat(cdf: CellDataFormat, workbook: XSSFWorkbook, cellStyle: XSSFCellStyle) {
    cdf.formatString.foreach(formatString => {
      val format = dataFormatCache.getOrElseUpdate(workbook, workbook.createDataFormat())
      val formatIndex = format.getFormat(formatString)
      cellStyle.setDataFormat(formatIndex)
    })
  }

  private def convertCellFill(cf: CellFill): FillPatternType = {
    import CellFill._

    cf match {
      case None => FillPatternType.NO_FILL
      case Solid => FillPatternType.SOLID_FOREGROUND
      case Pattern.AltBars => FillPatternType.ALT_BARS
      case Pattern.BigSpots => FillPatternType.BIG_SPOTS
      case Pattern.Bricks => FillPatternType.BRICKS
      case Pattern.Diamonds => FillPatternType.DIAMONDS
      case Pattern.Squares => FillPatternType.SQUARES
      case Pattern.Dots.Fine => FillPatternType.FINE_DOTS
      case Pattern.Dots.Least => FillPatternType.LEAST_DOTS
      case Pattern.Dots.Less => FillPatternType.LESS_DOTS
      case Pattern.Dots.Sparse => FillPatternType.SPARSE_DOTS
      case Pattern.Diagonals.ThickBackward => FillPatternType.THICK_BACKWARD_DIAG
      case Pattern.Diagonals.ThickForward => FillPatternType.THICK_FORWARD_DIAG
      case Pattern.Diagonals.ThinBackward => FillPatternType.THIN_BACKWARD_DIAG
      case Pattern.Diagonals.ThinForward => FillPatternType.THIN_FORWARD_DIAG
      case Pattern.Bands.ThickHorizontal => FillPatternType.THICK_HORZ_BANDS
      case Pattern.Bands.ThickVertical => FillPatternType.THICK_VERT_BANDS
      case Pattern.Bands.ThinHorizontal => FillPatternType.THIN_HORZ_BANDS
      case Pattern.Bands.ThinVertical => FillPatternType.THIN_VERT_BANDS
    }
  }

  private def convertCellStyle(cs: CellStyle, workbook: XSSFWorkbook): XSSFCellStyle =
    getCachedOrUpdate(cellStyleCache, cs, workbook) {
      val cellStyle = workbook.createCellStyle()
      cs.borders.foreach(b => convertCellBorders(b, cellStyle))
      cs.dataFormat.foreach(df => convertCellDataFormat(df, workbook, cellStyle))
      cs.font.foreach(f => f.convert(workbook))

      cs.fillPattern.foreach(fp => cellStyle.setFillPattern(convertCellFill(fp)))
      cs.fillBackgroundColor.foreach(c => cellStyle.setFillBackgroundColor(convertColor(c)))
      cs.fillForegroundColor.foreach(c => cellStyle.setFillForegroundColor(convertColor(c)))

      cs.horizontalAlignment.foreach(ha => cellStyle.setAlignment(ha.convertAsXlsx()))
      cs.verticalAlignment.foreach(va => cellStyle.setVerticalAlignment(va.convertAsXlsx()))

      cs.hidden.foreach(cellStyle.setHidden)
      cs.indention.foreach(cellStyle.setIndention)
      cs.locked.foreach(cellStyle.setLocked)
      cs.rotation.foreach(cellStyle.setRotation)
      cs.wrapText.foreach(cellStyle.setWrapText)
      cellStyle
    }

  private def convertHorizontalAlignment(horizontalAlignment: CellHorizontalAlignment): HorizontalAlignment = {
    import CellHorizontalAlignment._
    import HorizontalAlignment._

    horizontalAlignment match {
      case Center => CENTER
      case CenterSelection => CENTER_SELECTION
      case Disturbed => DISTRIBUTED
      case Fill => FILL
      case General => GENERAL
      case Justify => JUSTIFY
      case Left => LEFT
      case Right => RIGHT
      case Undefined =>
        throw new RuntimeException("Internal SPOIWO framework error. Undefined value is not allowed for explicit transformation!")
      case CellHorizontalAlignment(id) =>
        throw new Exception("Unsupported option for XLSX conversion with id=%d".format(id))
    }
  }

  private def convertVerticalAlignment(verticalAlignment: CellVerticalAlignment): VerticalAlignment = {
    import CellVerticalAlignment._
    import VerticalAlignment._

    verticalAlignment match {
      case Bottom => BOTTOM
      case Center => CENTER
      case Disturbed => DISTRIBUTED
      case Justify => JUSTIFY
      case Top => TOP
      case Undefined =>
        throw new RuntimeException("Internal SPOIWO framework error. Undefined value is not allowed for explicit transformation!")
      case CellVerticalAlignment(id) =>
        throw new RuntimeException("Unsupported option for XLSX conversion with id=%d".format(id))
    }
  }

  //================= Cache processing ====================
  def getCachedOrUpdate[K, V](cache: Cache[K, V], value: K, workbook: XSSFWorkbook)(newValue: => V): V = {
    val workbookCache = cache.getOrElseUpdate(workbook, collection.mutable.Map[K, V]())
    workbookCache.getOrElseUpdate(value, newValue)
  }

}
