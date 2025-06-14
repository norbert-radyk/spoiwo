package spoiwo.natures.xlsx

import java.time.format.DateTimeFormatter
import java.time.{LocalDate, ZoneId}
import java.util.{Calendar, Date}
import Model2XlsxEnumConversions._
import org.apache.poi.ss.usermodel
import org.apache.poi.ss.usermodel.{PageMargin, PaneType}
import org.apache.poi.ss.util.{CellAddress, CellRangeAddress}
import org.apache.poi.xssf.usermodel._
import spoiwo.model._
import spoiwo.model.enums._

trait BaseXlsx {

  private val FirstSupportedDate = LocalDate.of(1904, 1, 1)
  private val LastSupportedDate = LocalDate.of(9999, 12, 31)

  protected[natures] def mergeStyle(
      cell: Cell,
      rowStyle: Option[CellStyle],
      columnStyle: Option[CellStyle],
      sheetStyle: Option[CellStyle]
  ): Cell = {
    cell.styleInheritance match {
      case CellStyleInheritance.CellOnly =>
        cell
      case CellStyleInheritance.CellThenRow =>
        cell.withDefaultStyle(rowStyle)
      case CellStyleInheritance.CellThenColumn =>
        cell.withDefaultStyle(columnStyle)
      case CellStyleInheritance.CellThenSheet =>
        cell.withDefaultStyle(sheetStyle)
      case CellStyleInheritance.CellThenColumnThenRow =>
        cell.withDefaultStyle(columnStyle).withDefaultStyle(rowStyle)
      case CellStyleInheritance.CellThenRowThenColumn =>
        cell.withDefaultStyle(rowStyle).withDefaultStyle(columnStyle)
      case CellStyleInheritance.CellThenRowThenSheet =>
        cell.withDefaultStyle(rowStyle).withDefaultStyle(sheetStyle)
      case CellStyleInheritance.CellThenColumnThenSheet =>
        cell.withDefaultStyle(columnStyle).withDefaultStyle(sheetStyle)
      case CellStyleInheritance.CellThenColumnThenRowThenSheet =>
        cell.withDefaultStyle(columnStyle).withDefaultStyle(rowStyle).withDefaultStyle(sheetStyle)
      case CellStyleInheritance.CellThenRowThenColumnThenSheet =>
        cell.withDefaultStyle(rowStyle).withDefaultStyle(columnStyle).withDefaultStyle(sheetStyle)
      case unexpected =>
        throw new IllegalArgumentException(
          s"Unable to convert CellStyleInheritance=$unexpected to XLSX - unsupported enum!"
        )
    }
  }

  protected[natures] def convertColumn(c: Column, sheet: usermodel.Sheet): Unit = {
    val i = c.index.getOrElse(
      throw new IllegalArgumentException(
        "Undefined column index! " +
          "Something went terribly wrong as it should have been derived if not specified explicitly!"
      )
    )

    c.autoSized.foreach(as => if (as) sheet.autoSizeColumn(i))
    c.break.foreach(b => if (b) sheet.setColumnBreak(i))
    c.groupCollapsed.foreach(gc => sheet.setColumnGroupCollapsed(i, gc))
    c.hidden.foreach(h => sheet.setColumnHidden(i, h))
    c.width.foreach(w => sheet.setColumnWidth(i, w.toUnits))
  }

  protected[natures] def convertCellBorders(borders: CellBorders, style: XSSFCellStyle): Unit = {
    borders.leftStyle.foreach(s => style.setBorderLeft(convertBorderStyle(s)))
    borders.leftColor.foreach(c => style.setLeftBorderColor(convertColor(c)))
    borders.bottomStyle.foreach(s => style.setBorderBottom(convertBorderStyle(s)))
    borders.bottomColor.foreach(c => style.setBottomBorderColor(convertColor(c)))
    borders.rightStyle.foreach(s => style.setBorderRight(convertBorderStyle(s)))
    borders.rightColor.foreach(c => style.setRightBorderColor(convertColor(c)))
    borders.topStyle.foreach(s => style.setBorderTop(convertBorderStyle(s)))
    borders.topColor.foreach(c => style.setTopBorderColor(convertColor(c)))
  }

  protected[natures] def convertColor(color: Color): XSSFColor = new XSSFColor(
    Array[Byte](color.r.toByte, color.g.toByte, color.b.toByte),
    new DefaultIndexedColorMap()
  )

  protected[natures] def setHyperLinkCell(cell: usermodel.Cell, value: HyperLink, row: usermodel.Row): Unit = {
    val xlsxHyperlinkType = convertHyperlinkType(value.linkType)
    val link = row.getSheet.getWorkbook.getCreationHelper.createHyperlink(xlsxHyperlinkType)
    link.setAddress(value.address)
    cell.setCellValue(value.text)
    cell.setHyperlink(link)
  }

  protected[natures] def convertCellRange(cr: CellRange): CellRangeAddress =
    new org.apache.poi.ss.util.CellRangeAddress(cr.rowRange._1, cr.rowRange._2, cr.columnRange._1, cr.columnRange._2)

  protected[natures] def convertFont(f: Font, font: XSSFFont): XSSFFont = {
    f.bold.foreach(font.setBold)
    f.charSet.foreach(charSet => font.setCharSet(convertCharset(charSet).getNativeId))
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

  protected[natures] def validateCells(r: Row): Unit = {
    val indexedCells = r.cells.filter(_.index.isDefined)
    val contextCells = r.cells.filter(_.index.isEmpty)

    if (indexedCells.nonEmpty && contextCells.nonEmpty)
      throw new IllegalArgumentException("It is not allowed to mix cells with and without index within a single row!")

    val distinctIndexes = indexedCells.map(_.index).toSet.flatten
    if (indexedCells.size != distinctIndexes.size)
      throw new IllegalArgumentException("It is not allowed to have cells with duplicate index within a single row!")
  }

  protected[natures] def validateRows(s: Sheet): Unit = {
    val indexedRows = s.rows.filter(_.index.isDefined)
    val contextRows = s.rows.filter(_.index.isEmpty)

    if (indexedRows.nonEmpty && contextRows.nonEmpty) {
      throw new IllegalArgumentException("It is not allowed to mix rows with and without index within a single sheet!")
    }

    val distinctIndexes = indexedRows.flatMap(_.index).toSet
    if (indexedRows.size != distinctIndexes.size)
      throw new IllegalArgumentException("It is not allowed to have rows with duplicate index within a single sheet!")
  }

  protected[natures] def updateColumnsWithIndexes(s: Sheet): List[Column] = {
    val currentColumnIndexes = s.columns.flatMap(_.index).toSet
    if (currentColumnIndexes.isEmpty) {
      s.columns.zipWithIndex.map {
        case (column, index) => column.withIndex(index)
      }
    } else if (currentColumnIndexes.size == s.columns.size) {
      s.columns
    } else {
      throw new IllegalArgumentException(
        "When explicitly specifying column index you are required to provide it " +
          "uniquely for all columns in this sheet definition!"
      )
    }
  }

  protected[natures] def convertSheetProperties(sp: SheetProperties, sheet: usermodel.Sheet): Unit = {
    sp.autoFilter.foreach(autoFilterRange => sheet.setAutoFilter(convertCellRange(autoFilterRange)))
    sp.activeCell.foreach(stringReference => sheet.setActiveCell(new CellAddress(stringReference)))
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
      case CellRange((startRow, endRow), (startColumn, endColumn)) =>
        val workbook = sheet.getWorkbook
        val sheetIndex = workbook.getNumberOfSheets - 1
        workbook.setPrintArea(sheetIndex, startColumn, endColumn, startRow, endRow)
    }
    sp.printGridLines.foreach(sheet.setPrintGridlines)
    sp.rightToLeft.foreach(sheet.setRightToLeft)
    sp.rowSumsBelow.foreach(sheet.setRowSumsBelow)
    sp.rowSumsRight.foreach(sheet.setRowSumsRight)
    sp.selected.foreach(sheet.setSelected)
    sp.tabColor.foreach(c => setTabColor(sheet, convertColor(c)))
    sp.virtuallyCenter.foreach(sheet.setVerticallyCenter)
    sp.zoom.foreach(sheet.setZoom)
  }

  protected[natures] def convertSheetLocking(sl: SheetLocking, sheet: XSSFSheet): Unit = {
    sheet.lockAutoFilter(sl.lockedAutoFilter)
    sheet.lockDeleteColumns(sl.lockedDeleteColumns)
    sheet.lockDeleteRows(sl.lockedDeleteRows)
    sheet.lockFormatCells(sl.lockedFormatCells)
    sheet.lockFormatColumns(sl.lockedFormatColumns)
    sheet.lockFormatRows(sl.lockedFormatRows)
    sheet.lockInsertColumns(sl.lockedInsertColumns)
    sheet.lockInsertHyperlinks(sl.lockedInsertHyperlinks)
    sheet.lockInsertRows(sl.lockedInsertRows)
    sheet.lockPivotTables(sl.lockedPivotTables)
    sheet.lockSort(sl.lockedSort)
    sheet.lockObjects(sl.lockedObjects)
    sheet.lockScenarios(sl.lockedScenarios)
    sheet.lockSelectLockedCells(sl.lockedSelectLockedCells)
    sheet.lockSelectUnlockedCells(sl.lockedSelectUnlockedCells)
    sheet.enableLocking()
  }

  protected[natures] def setTabColor(sheet: usermodel.Sheet, XSSFColor: XSSFColor): Unit

  protected[natures] def convertPrintSetup(printSetup: PrintSetup, sheet: usermodel.Sheet): Unit = {
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
      printSetup.pageStart.foreach(ps.setPageStart)
      printSetup.scale.foreach(ps.setScale)
      printSetup.usePage.foreach(ps.setUsePage)
      printSetup.validSettings.foreach(ps.setValidSettings)
      printSetup.vResolution.foreach(ps.setVResolution)

      additionalPrintSetup(printSetup, ps)
    }
  }

  protected[natures] def additionalPrintSetup(printSetup: PrintSetup, sheetPs: usermodel.PrintSetup): Unit

  protected[natures] def convertRowRange(rr: RowRange): CellRangeAddress =
    CellRangeAddress.valueOf("%d:%d".format(rr.firstRowIndex, rr.lastRowIndex))

  protected[natures] def convertColumnRange(cr: ColumnRange): CellRangeAddress =
    CellRangeAddress.valueOf("%s:%s".format(cr.firstColumnName, cr.lastColumnName))

  protected[natures] def setDateCell(c: Cell, cell: usermodel.Cell, value: Date): Unit = {
    val dateStyle = DateTimeFormatter.ofPattern(c.format.getOrElse("yyyy-MM-dd"))
    val date = value.toInstant.atZone(ZoneId.systemDefault())
    if (date.toLocalDate.isBefore(FirstSupportedDate) || date.toLocalDate.isAfter(LastSupportedDate)) {
      cell.setCellValue(date.format(dateStyle))
    } else {
      cell.setCellValue(value)
    }
  }

  protected[natures] def setCalendarCell(c: Cell, cell: usermodel.Cell, value: Calendar): Unit = {
    val dateStyle = DateTimeFormatter.ofPattern(c.format.getOrElse("yyyy-MM-dd"))
    val date = value.toInstant.atZone(ZoneId.systemDefault())
    if (date.toLocalDate.isBefore(FirstSupportedDate) || date.toLocalDate.isAfter(LastSupportedDate)) {
      cell.setCellValue(date.format(dateStyle))
    } else {
      cell.setCellValue(value)
    }
  }

  protected[natures] def convertPaneType(pane: Pane): PaneType = pane match {
    case Pane.LowerLeftPane  => PaneType.LOWER_LEFT
    case Pane.LowerRightPane => PaneType.LOWER_RIGHT
    case Pane.UpperLeftPane  => PaneType.UPPER_LEFT
    case Pane.UpperRightPane => PaneType.UPPER_RIGHT
    case unexpected =>
      throw new IllegalArgumentException(s"Unable to convert Pane=$unexpected to XLSX - unsupported enum!")
  }

  protected[natures] def convertPaneAction(paneAction: PaneAction, sheet: usermodel.Sheet): Unit = {
    paneAction match {
      case NoSplitOrFreeze() =>
        sheet.createFreezePane(0, 0)
      case FreezePane(columnSplit, rowSplit, leftMostColumn, topRow) =>
        sheet.createFreezePane(columnSplit, rowSplit, leftMostColumn, topRow)
      case SplitPane(xSplitPosition, ySplitPosition, leftMostColumn, topRow, activePane) =>
        sheet.createSplitPane(xSplitPosition, ySplitPosition, leftMostColumn, topRow, convertPaneType(activePane))
    }
  }

  protected[natures] def convertMargins(margins: Margins, sheet: usermodel.Sheet): Unit = {
    margins.top.foreach(topMargin => sheet.setMargin(PageMargin.TOP, topMargin))
    margins.bottom.foreach(bottomMargin => sheet.setMargin(PageMargin.BOTTOM, bottomMargin))
    margins.right.foreach(rightMargin => sheet.setMargin(PageMargin.RIGHT, rightMargin))
    margins.left.foreach(leftMargin => sheet.setMargin(PageMargin.LEFT, leftMargin))
    margins.header.foreach(headerMargin => sheet.setMargin(PageMargin.HEADER, headerMargin))
    margins.footer.foreach(footerMargin => sheet.setMargin(PageMargin.FOOTER, footerMargin))
  }

}
