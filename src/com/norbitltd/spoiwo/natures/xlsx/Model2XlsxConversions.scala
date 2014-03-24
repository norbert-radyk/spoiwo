package com.norbitltd.spoiwo.natures.xlsx

import org.apache.poi.ss.usermodel.{HorizontalAlignment, VerticalAlignment}
import com.norbitltd.spoiwo.model._
import org.apache.poi.xssf.usermodel._
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.usermodel.Sheet._
import com.norbitltd.spoiwo.model.SplitPane
import com.norbitltd.spoiwo.model.NoSplitOrFreeze
import java.io.FileOutputStream
import Model2XlsxEnumConversions._
import com.norbitltd.spoiwo.model.enums.{Pane, CellFill}

object Model2XlsxConversions {

  private type Cache[K, V] = collection.mutable.Map[XSSFWorkbook, collection.mutable.Map[K, V]]

  private def Cache[K, V]() = collection.mutable.Map[XSSFWorkbook, collection.mutable.Map[K, V]]()

  private lazy val cellStyleCache = Cache[CellStyle, XSSFCellStyle]()
  private lazy val dataFormatCache = collection.mutable.Map[XSSFWorkbook, XSSFDataFormat]()
  private lazy val fontCache = Cache[Font, XSSFFont]()

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
    def convertAsXlsx(workbook: XSSFWorkbook): XSSFCellStyle = convertCellStyle(cs, workbook)
  }

  implicit class XlsxFont(f: Font) {
    def convertAsXlsx(cell: XSSFCell): XSSFFont = convertAsXlsx(cell.getRow)
    def convertAsXlsx(row: XSSFRow): XSSFFont = convertAsXlsx(row.getSheet)
    def convertAsXlsx(sheet: XSSFSheet): XSSFFont = convertAsXlsx(sheet.getWorkbook)
    def convertAsXlsx(workbook: XSSFWorkbook): XSSFFont = convertFont(f, workbook)
  }

  implicit class XlsxHorizontalAlignment(ha: CellHorizontalAlignment) {
    def convertAsXlsx() = convertHorizontalAlignment(ha)
  }

  implicit class XlsxSheet(s: Sheet) {
    def convertAsXlsx(workbook: XSSFWorkbook) = convertSheet(s, workbook)

    def convertAsXlsx() = Workbook(s).convertAsXlsx()

    def saveAsXlsx(fileName: String) {
      Workbook(s).saveAsXlsx(fileName)
    }
  }

  implicit class XlsxVerticalAlignment(va: CellVerticalAlignment) {
    def convertAsXlsx() = convertVerticalAlignment(va)
  }

  implicit class XlsxWorkbook(workbook : Workbook) {
    def convertAsXlsx() = convertWorkbook(workbook)

    def saveAsXlsx(fileName : String) = {
      val stream = new FileOutputStream(fileName)
      try {
        val workbook = convertAsXlsx()
        workbook.write(stream)
      } finally {
        stream.flush()
        stream.close()
      }
    }
  }

  private def convertColor(color: Color): XSSFColor = new XSSFColor(
    Array[Byte](color.r.toByte, color.g.toByte, color.b.toByte)
  )

  private[xlsx] def convertCell(c : Cell, row : XSSFRow): XSSFCell = {
    val cellNumber = c.index.getOrElse(if (row.getLastCellNum < 0) 0 else row.getLastCellNum)
    val cell = row.createCell(cellNumber)
    c.style.foreach(s => cell.setCellStyle(convertCellStyle(s, cell.getRow.getSheet.getWorkbook)))
    c match {
      case StringCell(value, _, _) => cell.setCellValue(value)
      case FormulaCell(formula, _, _) => cell.setCellFormula(formula)
      case NumericCell(value, _, _) => cell.setCellValue(value)
      case BooleanCell(value, _, _) => cell.setCellValue(value)
      case DateCell(value, _, _) => cell.setCellValue(value)
      case CalendarCell(value, _, _) => cell.setCellValue(value)
    }
    cell
  }

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

  private def convertCellRange(cr: CellRange): CellRangeAddress =
    new org.apache.poi.ss.util.CellRangeAddress(cr.rowRange._1, cr.rowRange._2, cr.columnRange._1, cr.columnRange._2)

  private[xlsx] def convertCellStyle(cs: CellStyle, workbook: XSSFWorkbook): XSSFCellStyle =
    getCachedOrUpdate(cellStyleCache, cs, workbook) {
      val cellStyle = workbook.createCellStyle()
      cs.borders.foreach(b => convertCellBorders(b, cellStyle))
      cs.dataFormat.foreach(df => convertCellDataFormat(df, workbook, cellStyle))
      cs.font.foreach(f => cellStyle.setFont(convertFont(f, workbook)))
      cs.fillPattern.foreach(fp => cellStyle.setFillPattern(convertCellFill(fp)))
      cs.fillBackgroundColor.foreach(c => cellStyle.setFillBackgroundColor(convertColor(c)))
      cs.fillForegroundColor.foreach(c => cellStyle.setFillForegroundColor(convertColor(c)))

      cs.horizontalAlignment.foreach(ha => cellStyle.setAlignment(convertHorizontalAlignment(ha)))
      cs.verticalAlignment.foreach(va => cellStyle.setVerticalAlignment(convertVerticalAlignment(va)))

      cs.hidden.foreach(cellStyle.setHidden)
      cs.indention.foreach(cellStyle.setIndention)
      cs.locked.foreach(cellStyle.setLocked)
      cs.rotation.foreach(cellStyle.setRotation)
      cs.wrapText.foreach(cellStyle.setWrapText)
      cellStyle
    }


  private[xlsx] def convertColumn(c : Column, sheet: XSSFSheet) {
    val i = c.index.getOrElse(throw new IllegalArgumentException("Undefined column index! " +
      "Something went terribly wrong as it should have been derived if not specified explicitly!"))

    c.autoSized.foreach(as => sheet.autoSizeColumn(i))
    c.break.foreach(b => sheet.setColumnBreak(i))
    c.groupCollapsed.foreach(gc => sheet.setColumnGroupCollapsed(i, gc))
    c.hidden.foreach(h => sheet.setColumnHidden(i, h))
    c.style.foreach(s => sheet.setDefaultColumnStyle(i, convertCellStyle(s, sheet.getWorkbook)))
    c.width.foreach(w => sheet.setColumnWidth(i, w))
  }

  private def convertColumnRange(cr: ColumnRange) = CellRangeAddress.valueOf("%s:%s".format(cr.firstColumnName, cr.lastColumnName))

  private def convertFooter(f: Footer, sheet: XSSFSheet) {
    f.left.foreach(sheet.getFooter.setLeft)
    f.center.foreach(sheet.getFooter.setCenter)
    f.right.foreach(sheet.getFooter.setRight)

    f.firstLeft.foreach(sheet.getFirstFooter.setLeft)
    f.firstCenter.foreach(sheet.getFirstFooter.setCenter)
    f.firstRight.foreach(sheet.getFirstFooter.setRight)

    f.oddLeft.foreach(sheet.getOddFooter.setLeft)
    f.oddCenter.foreach(sheet.getOddFooter.setCenter)
    f.oddRight.foreach(sheet.getOddFooter.setRight)

    f.evenLeft.foreach(sheet.getEvenFooter.setLeft)
    f.evenCenter.foreach(sheet.getEvenFooter.setCenter)
    f.evenRight.foreach(sheet.getEvenFooter.setRight)
  }

  private[xlsx] def convertFont(f : Font, workbook : XSSFWorkbook) : XSSFFont =
    getCachedOrUpdate(fontCache, f, workbook) {
      val font = workbook.createFont()
      f.bold.foreach(font.setBold)
      f.charSet.foreach(charSet => font.setCharSet(convertCharset(charSet)))
      f.color.foreach(c => font.setColor(convertColor(c)))
      f.family.foreach(family => font.setFamily(convertFontFamily(family)))
      f.height.foreach(height => font.setFontHeightInPoints(height.toPoints))
      f.italic.foreach(font.setItalic)
      f.scheme.foreach(scheme => font.setScheme(convertFontScheme(scheme)))
      f.fontName.foreach(font.setFontName)
      f.strikeout.foreach(font.setStrikeout)
      f.typeOffset.foreach(offset => font.setTypeOffset(convertTypeOffset(offset)))
      f.underline.foreach(underline => font.setUnderline(convertUnderline(underline)))
      font
    }


  private def convertHeader(h: Header, sheet: XSSFSheet) {
    h.left.foreach(sheet.getHeader.setLeft)
    h.center.foreach(sheet.getHeader.setCenter)
    h.right.foreach(sheet.getHeader.setRight)

    h.firstLeft.foreach(sheet.getFirstHeader.setLeft)
    h.firstCenter.foreach(sheet.getFirstHeader.setCenter)
    h.firstRight.foreach(sheet.getFirstHeader.setRight)

    h.oddLeft.foreach(sheet.getOddHeader.setLeft)
    h.oddCenter.foreach(sheet.getOddHeader.setCenter)
    h.oddRight.foreach(sheet.getOddHeader.setRight)

    h.evenLeft.foreach(sheet.getEvenHeader.setLeft)
    h.evenCenter.foreach(sheet.getEvenHeader.setCenter)
    h.evenRight.foreach(sheet.getEvenHeader.setRight)
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

  private def convertMargins(margins: Margins, sheet: XSSFSheet) {
    margins.top.foreach(topMargin => sheet.setMargin(TopMargin, topMargin))
    margins.bottom.foreach(bottomMargin => sheet.setMargin(BottomMargin, bottomMargin))
    margins.right.foreach(rightMargin => sheet.setMargin(RightMargin, rightMargin))
    margins.left.foreach(leftMargin => sheet.setMargin(LeftMargin, leftMargin))
    margins.header.foreach(headerMargin => sheet.setMargin(HeaderMargin, headerMargin))
    margins.footer.foreach(footerMargin => sheet.setMargin(FooterMargin, footerMargin))
  }

  private def convertPane(pane: Pane): Int = pane match {
    case Pane.LowerLeftPane => org.apache.poi.ss.usermodel.Sheet.PANE_LOWER_LEFT
    case Pane.LowerRightPane => org.apache.poi.ss.usermodel.Sheet.PANE_LOWER_RIGHT
    case Pane.UpperLeftPane => org.apache.poi.ss.usermodel.Sheet.PANE_UPPER_LEFT
    case Pane.UpperRightPane => org.apache.poi.ss.usermodel.Sheet.PANE_UPPER_RIGHT
  }

  private def convertPaneAction(paneAction: PaneAction, sheet: XSSFSheet) {
    paneAction match {
      case NoSplitOrFreeze() =>
        sheet.createFreezePane(0, 0)
      case FreezePane(columnSplit, rowSplit, leftMostColumn, topRow) =>
        sheet.createFreezePane(columnSplit, rowSplit, leftMostColumn, topRow)
      case SplitPane(xSplitPosition, ySplitPosition, leftMostColumn, topRow, activePane) =>
        sheet.createSplitPane(xSplitPosition, ySplitPosition, leftMostColumn, topRow, convertPane(activePane))
    }
  }

  private[xlsx] def convertRow(r : com.norbitltd.spoiwo.model.Row, sheet: XSSFSheet): XSSFRow = {
    validateCells(r)
    val indexNumber = r.index.getOrElse(if (sheet.rowIterator().hasNext) sheet.getLastRowNum + 1 else 0)
    val row = sheet.createRow(indexNumber)

    r.cells.foreach(cell => convertCell(cell, row))
    r.height.foreach(h => row.setHeightInPoints(h.toPoints))
    r.style.foreach(s => row.setRowStyle(convertCellStyle(s, row.getSheet.getWorkbook)))
    r.hidden.foreach(row.setZeroHeight)
    row
  }

  private def validateCells(r : Row) {
    val indexedCells = r.cells.filter(_.index.isDefined)
    val contextCells = r.cells.filter(_.index.isEmpty)

    if(indexedCells.size > 0 && contextCells.size > 0)
      throw new IllegalArgumentException("It is not allowed to mix cells with and without index within a single row!")

    val distinctIndexes = indexedCells.map(_.index).toSet.flatten
    if(indexedCells.size != distinctIndexes.size)
      throw new IllegalArgumentException("It is not allowed to have cells with duplicate index within a single row!")
  }

  private def convertSheet(s: Sheet, workbook: XSSFWorkbook): XSSFSheet = {
    val sheetName = s.name.getOrElse("Sheet " + (workbook.getNumberOfSheets + 1))
    val sheet = workbook.createSheet(sheetName)

    updateColumnsWithIndexes(s).foreach(column => convertColumn(column, sheet))
    s.rows.foreach(row => convertRow(row, sheet))
    s.mergedRegions.foreach(mergedRegion => sheet.addMergedRegion(convertCellRange(mergedRegion)))

    s.printSetup.foreach(ps => convertPrintSetup(ps, sheet))
    s.header.foreach(h => convertHeader(h, sheet))
    s.footer.foreach(f => convertFooter(f, sheet))

    s.properties.foreach(sp => convertSheetProperties(sp, sheet))
    s.margins.foreach(m => convertMargins(m, sheet))
    s.paneAction.foreach(pa => convertPaneAction(pa, sheet))
    s.repeatingRows.foreach(rr => sheet.setRepeatingRows(convertRowRange(rr)))
    s.repeatingColumns.foreach(rc => sheet.setRepeatingColumns(convertColumnRange(rc)))
    sheet
  }

  private def updateColumnsWithIndexes(s: Sheet): List[Column] = {
    val currentColumnIndexes = s.columns.map(_.index).flatten.toSet
    if (currentColumnIndexes.isEmpty) {
      s.columns.zipWithIndex.map {
        case (column, index) => column.withIndex(index)
      }
    } else if (currentColumnIndexes.size == s.columns.size) {
      s.columns
    } else {
      throw new IllegalArgumentException(
        "When explicitly specifying column index you are required to provide it " +
          "uniquely for all columns in this sheet definition!")
    }
  }

  private[xlsx] def convertSheetProperties(sp: SheetProperties, sheet: XSSFSheet) {
    sp.autoFilter.foreach(autoFilterRange => sheet.setAutoFilter(convertCellRange(autoFilterRange)))
    sp.activeCell.foreach(sheet.setActiveCell)
    sp.autoBreaks.foreach(sheet.setAutobreaks)
    sp.defaultColumnWidth.foreach(sheet.setDefaultColumnWidth)
    sp.defaultRowHeight.foreach(height => sheet.setDefaultRowHeightInPoints(height.toPoints))
    sp.displayFormulas.foreach(sheet.setDisplayFormulas)
    sp.displayGridLines.foreach(sheet.setDisplayGridlines)
    sp.displayGuts.foreach(sheet.setDisplayGuts)
    sp.displayRowColHeadings.foreach(sheet.setDisplayRowColHeadings)
    sp.displayZeros.foreach(sheet.setDisplayZeros)
    sp.fitToPage.foreach(sheet.setFitToPage)
    sp.forceFormulaRecalculation.foreach(sheet.setForceFormulaRecalculation)
    sp.horizontallyCenter.foreach(sheet.setHorizontallyCenter)
    sp.printArea.foreach {
      case CellRange((startRow, endRow), (startColumn, endColumn)) => {
        val workbook = sheet.getWorkbook
        val sheetIndex = workbook.getNumberOfSheets - 1
        workbook.setPrintArea(sheetIndex, startColumn, endColumn, startRow, endRow)
      }
    }
    sp.printGridLines.foreach(sheet.setPrintGridlines)
    sp.rightToLeft.foreach(sheet.setRightToLeft)
    sp.rowSumsBelow.foreach(sheet.setRowSumsBelow)
    sp.rowSumsRight.foreach(sheet.setRowSumsRight)
    sp.selected.foreach(sheet.setSelected)
    sp.tabColor.foreach(sheet.setTabColor)
    sp.virtuallyCenter.foreach(sheet.setVerticallyCenter)
    sp.zoom.foreach(sheet.setZoom)
  }

  private def convertPrintSetup(printSetup: PrintSetup, sheet: XSSFSheet) {
    if (printSetup != PrintSetup.Default) {
      val ps = sheet.getPrintSetup
      printSetup.copies.foreach(ps.setCopies)
      printSetup.draft.foreach(ps.setDraft)
      printSetup.fitHeight.foreach(ps.setFitHeight)
      printSetup.fitWidth.foreach(ps.setFitWidth)
      printSetup.footerMargin.foreach(ps.setFooterMargin)
      printSetup.headerMargin.foreach(ps.setHeaderMargin)
      printSetup.hResolution.foreach(ps.setHResolution)
      printSetup.landscape.foreach(ps.setLandscape)
      printSetup.leftToRight.foreach(ps.setLeftToRight)
      printSetup.noColor.foreach(ps.setNoColor)
      printSetup.noOrientation.foreach(ps.setNoOrientation)
      printSetup.pageOrder.foreach(ps.setPageOrder)
      printSetup.pageStart.foreach(ps.setPageStart)
      printSetup.paperSize.foreach(ps.setPaperSize)
      printSetup.scale.foreach(ps.setScale)
      printSetup.usePage.foreach(ps.setUsePage)
      printSetup.validSettings.foreach(ps.setValidSettings)
      printSetup.vResolution.foreach(ps.setVResolution)
    }
  }

  private def convertRowRange(rr: RowRange) = CellRangeAddress.valueOf("%d:%d".format(rr.firstRowIndex, rr.lastRowIndex))

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

  private def convertWorkbook(wb : Workbook): XSSFWorkbook = {
    val workbook = new XSSFWorkbook()
    wb.sheets.foreach(sheet => convertSheet(sheet, workbook))

    //Parameters
    wb.activeSheet.foreach(workbook.setActiveSheet)
    wb.firstVisibleTab.foreach(workbook.setFirstVisibleTab)
    wb.forceFormulaRecalculation.foreach(workbook.setForceFormulaRecalculation)
    wb.hidden.foreach(workbook.setHidden)
    wb.missingCellPolicy.foreach(mcp => workbook.setMissingCellPolicy(convertMissingCellPolicy(mcp)))
    wb.selectedTab.foreach(workbook.setSelectedTab)
    workbook
  }

  //================= Cache processing ====================
  private def getCachedOrUpdate[K, V](cache: Cache[K, V], value: K, workbook: XSSFWorkbook)(newValue: => V): V = {
    val workbookCache = cache.getOrElseUpdate(workbook, collection.mutable.Map[K, V]())
    workbookCache.getOrElseUpdate(value, newValue)
  }

}
