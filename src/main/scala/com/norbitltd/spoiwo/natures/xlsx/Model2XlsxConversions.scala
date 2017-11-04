package com.norbitltd.spoiwo.natures.xlsx

import java.io.FileOutputStream
import java.util.{Calendar, Date}

import com.norbitltd.spoiwo.model.{BooleanCell, CalendarCell, DateCell, FormulaCell, NoSplitOrFreeze, NumericCell, SplitPane, StringCell, _}
import com.norbitltd.spoiwo.model.enums._
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxEnumConversions._
import org.apache.poi.common.usermodel.HyperlinkType
import org.apache.poi.ss.usermodel
import org.apache.poi.ss.usermodel.{BorderStyle, FillPatternType, HorizontalAlignment, VerticalAlignment}
import org.apache.poi.ss.util.{CellAddress, CellRangeAddress}
import org.apache.poi.xssf.usermodel._
import org.openxmlformats.schemas.spreadsheetml.x2006.main.{CTTable, CTTableColumns, CTTableStyleInfo}
import org.joda.time.{LocalDate, LocalDateTime}

object Model2XlsxConversions {

  private type Cache[K, V] = collection.mutable.Map[XSSFWorkbook, collection.mutable.Map[K, V]]
  private lazy val cellStyleCache = Cache[CellStyle, XSSFCellStyle]()
  private lazy val dataFormatCache = collection.mutable.Map[XSSFWorkbook, XSSFDataFormat]()
  private lazy val fontCache = Cache[Font, XSSFFont]()
  private val FirstSupportedDate = new LocalDate(1904, 1, 1)
  private val LastSupportedDate = new LocalDate(9999, 12, 31)

  private def Cache[K, V]() = collection.mutable.Map[XSSFWorkbook, collection.mutable.Map[K, V]]()

  private def convertColor(color: Color): XSSFColor = new XSSFColor(
    Array[Byte](color.r.toByte, color.g.toByte, color.b.toByte),
    new DefaultIndexedColorMap()
  )

  private def mergeStyle(cell: Cell,
                         rowStyle: Option[CellStyle],
                         columnStyle: Option[CellStyle],
                         sheetStyle: Option[CellStyle]): Cell = {
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
        throw new IllegalArgumentException(s"Unable to convert CellStyleInheritance=$unexpected to XLSX - unsupported enum!")
    }
  }

  private[xlsx] def convertCell(modelSheet: Sheet, modelColumns: Map[Int, Column], modelRow: Row, c: Cell, row: XSSFRow): XSSFCell = {
    val cellNumber = c.index.getOrElse(if (row.getLastCellNum < 0) 0 else row.getLastCellNum)
    val cell = row.createCell(cellNumber)

    val cellWithStyle = mergeStyle(c, modelRow.style, modelColumns.get(cellNumber).flatMap(_.style), modelSheet.style)
    cellWithStyle.style.foreach(s => cell.setCellStyle(convertCellStyle(s, cell.getRow.getSheet.getWorkbook)))

    c match {
      case StringCell(value, _, _, _) => cell.setCellValue(value)
      case FormulaCell(formula, _, _, _) => cell.setCellFormula(formula)
      case NumericCell(value, _, _, _) => cell.setCellValue(value)
      case BooleanCell(value, _, _, _) => cell.setCellValue(value)
      case DateCell(value, _, _, _) => setDateCell(c, cell, value)
      case CalendarCell(value, _, _, _) => setCalendarCell(c, cell, value)
      case HyperLinkUrlCell(value, _, _, _) => setHyperLinkUrlCell(cell, value, row)
    }
    cell
  }

  private def setDateCell(c: Cell, cell: XSSFCell, value: Date) {
    val dateStyle = c.format.getOrElse("yyyy-MM-dd")
    val dateTime = LocalDateTime.fromDateFields(value)
    val date = dateTime.toLocalDate
    if (date.isBefore(FirstSupportedDate) || date.isAfter(LastSupportedDate)) {
      cell.setCellValue(dateTime.toString(dateStyle))
    } else {
      cell.setCellValue(value)
    }
  }

  private def setCalendarCell(c: Cell, cell: XSSFCell, value: Calendar) {
    val dateStyle = c.format.getOrElse("yyyy-MM-dd")
    val dateTime = LocalDateTime.fromCalendarFields(value)
    val date = dateTime.toLocalDate
    if (date.isBefore(FirstSupportedDate) || date.isAfter(LastSupportedDate)) {
      cell.setCellValue(dateTime.toString(dateStyle))
    } else {
      cell.setCellValue(value)
    }
  }
  private def setHyperLinkUrlCell(cell: XSSFCell, value: HyperLinkUrl, row: XSSFRow) {
    val link = row.getSheet.getWorkbook.getCreationHelper.createHyperlink(HyperlinkType.URL)
    link.setAddress(value.address)
    cell.setCellValue(value.text)
    cell.setHyperlink(link)

  }
  private[xlsx] def convertCellBorders(borders: CellBorders, style: XSSFCellStyle) {
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

  private[xlsx] def convertColumn(c: Column, sheet: XSSFSheet) {
    val i = c.index.getOrElse(throw new IllegalArgumentException("Undefined column index! " +
      "Something went terribly wrong as it should have been derived if not specified explicitly!"))

    c.autoSized.foreach(as => if (as) sheet.autoSizeColumn(i))
    c.break.foreach(b => if (b) sheet.setColumnBreak(i))
    c.groupCollapsed.foreach(gc => sheet.setColumnGroupCollapsed(i, gc))
    c.hidden.foreach(h => sheet.setColumnHidden(i, h))
    c.width.foreach(w => sheet.setColumnWidth(i, w.toUnits))
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

  private[xlsx] def convertFont(f: Font, workbook: XSSFWorkbook): XSSFFont =
    getCachedOrUpdate(fontCache, f, workbook) {
      val font = workbook.createFont()
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

  private def convertMargins(margins: Margins, sheet: XSSFSheet) {
    margins.top.foreach(topMargin => sheet.setMargin(usermodel.Sheet.TopMargin, topMargin))
    margins.bottom.foreach(bottomMargin => sheet.setMargin(usermodel.Sheet.BottomMargin, bottomMargin))
    margins.right.foreach(rightMargin => sheet.setMargin(usermodel.Sheet.RightMargin, rightMargin))
    margins.left.foreach(leftMargin => sheet.setMargin(usermodel.Sheet.LeftMargin, leftMargin))
    margins.header.foreach(headerMargin => sheet.setMargin(usermodel.Sheet.HeaderMargin, headerMargin))
    margins.footer.foreach(footerMargin => sheet.setMargin(usermodel.Sheet.FooterMargin, footerMargin))
  }

  private def convertPane(pane: Pane): Int = pane match {
    case Pane.LowerLeftPane => org.apache.poi.ss.usermodel.Sheet.PANE_LOWER_LEFT
    case Pane.LowerRightPane => org.apache.poi.ss.usermodel.Sheet.PANE_LOWER_RIGHT
    case Pane.UpperLeftPane => org.apache.poi.ss.usermodel.Sheet.PANE_UPPER_LEFT
    case Pane.UpperRightPane => org.apache.poi.ss.usermodel.Sheet.PANE_UPPER_RIGHT
    case unexpected =>
      throw new IllegalArgumentException(s"Unable to convert Pane=$unexpected to XLSX - unsupported enum!")
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

  private[xlsx] def convertRow(modelColumns: Map[Int, Column], modelRow: Row, modelSheet: Sheet, sheet: XSSFSheet): XSSFRow = {
    validateCells(modelRow)
    val indexNumber = modelRow.index.getOrElse(if (sheet.rowIterator().hasNext) sheet.getLastRowNum + 1 else 0)
    val row = sheet.createRow(indexNumber)
    modelRow.height.foreach(h => row.setHeightInPoints(h.toPoints))
    modelRow.style.foreach(s => row.setRowStyle(convertCellStyle(s, row.getSheet.getWorkbook)))
    modelRow.hidden.foreach(row.setZeroHeight)
    modelRow.cells.foreach(cell => convertCell(modelSheet, modelColumns, modelRow, cell, row))
    row
  }

  private def validateCells(r: Row) {
    val indexedCells = r.cells.filter(_.index.isDefined)
    val contextCells = r.cells.filter(_.index.isEmpty)

    if (indexedCells.nonEmpty && contextCells.nonEmpty)
      throw new IllegalArgumentException("It is not allowed to mix cells with and without index within a single row!")

    val distinctIndexes = indexedCells.map(_.index).toSet.flatten
    if (indexedCells.size != distinctIndexes.size)
      throw new IllegalArgumentException("It is not allowed to have cells with duplicate index within a single row!")
  }

  private[xlsx] def convertSheet(s: Sheet, workbook: XSSFWorkbook): XSSFSheet = {
    validateRows(s)
    validateTables(s)
    val sheetName = s.name.getOrElse("Sheet" + (workbook.getNumberOfSheets + 1))
    val sheet = workbook.createSheet(sheetName)
    val columns = updateColumnsWithIndexes(s)
    val columnsMap = columns.map(c => c.index.get -> c).toMap

    s.rows.foreach(row => convertRow(columnsMap, row, s, sheet))
    columns.foreach(c => convertColumn(c, sheet))
    s.mergedRegions.foreach(mergedRegion => sheet.addMergedRegion(convertCellRange(mergedRegion)))

    s.printSetup.foreach(ps => convertPrintSetup(ps, sheet))
    s.header.foreach(h => convertHeader(h, sheet))
    s.footer.foreach(f => convertFooter(f, sheet))

    s.properties.foreach(sp => convertSheetProperties(sp, sheet))
    s.margins.foreach(m => convertMargins(m, sheet))
    s.paneAction.foreach(pa => convertPaneAction(pa, sheet))
    s.repeatingRows.foreach(rr => sheet.setRepeatingRows(convertRowRange(rr)))
    s.repeatingColumns.foreach(rc => sheet.setRepeatingColumns(convertColumnRange(rc)))
    s.password.foreach(ps => sheet.protectSheet(ps))
    val tables = updateTablesWithIds(s)
    tables.foreach(tbl ⇒ convertTable(tbl, sheet))
    sheet
  }

  private def validateRows(s: Sheet) {
    val indexedRows = s.rows.filter(_.index.isDefined)
    val contextRows = s.rows.filter(_.index.isEmpty)

    if (indexedRows.nonEmpty && contextRows.nonEmpty) {
      throw new IllegalArgumentException("It is not allowed to mix rows with and without index within a single sheet!")
    }

    val distinctIndexes = indexedRows.flatMap(_.index).toSet
    if (indexedRows.size != distinctIndexes.size)
      throw new IllegalArgumentException("It is not allowed to have rows with duplicate index within a single sheet!")
  }

  private def updateColumnsWithIndexes(s: Sheet): List[Column] = {
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
          "uniquely for all columns in this sheet definition!")
    }
  }

  private[xlsx] def convertSheetProperties(sp: SheetProperties, sheet: XSSFSheet) {
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
    sp.tabColor.foreach(c => sheet.setTabColor(convertColor(c)))
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
      printSetup.pageOrder.foreach(po => ps.setPageOrder(convertPageOrder(po)))
      printSetup.pageStart.foreach(ps.setPageStart)
      printSetup.paperSize.foreach(px => ps.setPaperSize(convertPaperSize(px)))
      printSetup.scale.foreach(ps.setScale)
      printSetup.usePage.foreach(ps.setUsePage)
      printSetup.validSettings.foreach(ps.setValidSettings)
      printSetup.vResolution.foreach(ps.setVResolution)
    }
  }

  private def convertRowRange(rr: RowRange) = CellRangeAddress.valueOf("%d:%d".format(rr.firstRowIndex, rr.lastRowIndex))

  private[xlsx] def validateTables(modelSheet: Sheet): Unit = {
    val ids = modelSheet.tables.flatMap(_.id)
    if (ids.size != ids.toSet.size)
      throw new IllegalArgumentException("Specified table ids need to be unique.")
  }

  private[xlsx] def updateTablesWithIds(modelSheet: Sheet): List[Table] = {
    modelSheet.tables.zipWithIndex.map {
      case (table, index) ⇒ if (table.id.isDefined) table else table.withId(index + 1)
    }
  }

  private[xlsx] def convertTable(modelTable: Table, sheet: XSSFSheet): XSSFTable = {
    validateTableColumns(modelTable)

    val tableId = modelTable.id.getOrElse {
      throw new IllegalArgumentException("Undefined table id! " +
        "Something went terribly wrong as it should have been derived if not specified explicitly!")
    }

    val displayName = modelTable.displayName.getOrElse(s"Table$tableId")
    val name        = modelTable.name.getOrElse(s"ct_table_$tableId")

    val table = sheet.createTable()
    table.setDisplayName(displayName)
    val ctTable = table.getCTTable
    ctTable.setId(tableId)
    ctTable.setName(name)
    setTableReference(modelTable, ctTable)
    convertTableColumns(modelTable, ctTable)
    modelTable.style.foreach(convertTableStyle(_, ctTable))
    modelTable.enableAutoFilter.foreach(af ⇒ if (af) ctTable.addNewAutoFilter())
    table
  }

  private[xlsx] def validateTableColumns(modelTable: Table): Unit = {

    def insufficientColumnsDefined = {
      val (sCol, eCol)   = modelTable.cellRange.columnRange
      val neededColumns  = (eCol-sCol) + 1
      val definedColumns = modelTable.columns.size
      neededColumns != definedColumns
    }

    if (modelTable.columns.nonEmpty && insufficientColumnsDefined) {
      throw new IllegalArgumentException(
        "When explicitly specifying table columns you are required to provide as many " +
          "columns as in the cell range."
      )
    }
  }

  private[xlsx] def convertTableColumns(modelTable: Table, ctTable: CTTable): CTTableColumns = {

    def generateColumns = {
      val (sCol, eCol)   = modelTable.cellRange.columnRange
      val neededColumns  = (eCol-sCol) + 1
      (0 until neededColumns) map { index ⇒
        val columnId = index + 1
        TableColumn(
          name = s"TableColumn$columnId",
          id   = columnId.toLong
        )
      }
    }

    val modelColumns = if(modelTable.columns.nonEmpty) modelTable.columns else generateColumns
    val columns = ctTable.addNewTableColumns()
    columns.setCount(modelColumns.size)
    modelColumns.foreach { mc ⇒
      val column = columns.addNewTableColumn()
      column.setName(mc.name)
      column.setId(mc.id)
    }
    columns
  }

  private[xlsx] def convertTableStyle(modelTableStyle: TableStyle, ctTable: CTTable): CTTableStyleInfo = {
    val style = ctTable.addNewTableStyleInfo()
    style.setName(modelTableStyle.name.value)
    modelTableStyle.showColumnStripes.foreach(style.setShowColumnStripes)
    modelTableStyle.showRowStripes.foreach(style.setShowRowStripes)
    style
  }

  private[xlsx] def setTableReference(modelTable: Table, ctTable: CTTable): Unit = {
    val cellRangeAddress = convertCellRange(modelTable.cellRange)
    ctTable.setRef(cellRangeAddress.formatAsString())
  }

  private[xlsx] def convertWorkbook(wb: Workbook): XSSFWorkbook = {
    val workbook = new XSSFWorkbook()
    wb.sheets.foreach(sheet => convertSheet(sheet, workbook))

    //Parameters
    wb.activeSheet.foreach(workbook.setActiveSheet)
    wb.firstVisibleTab.foreach(workbook.setFirstVisibleTab)
    wb.forceFormulaRecalculation.foreach(workbook.setForceFormulaRecalculation)
    wb.hidden.foreach(workbook.setHidden)
    wb.missingCellPolicy.foreach(mcp => workbook.setMissingCellPolicy(convertMissingCellPolicy(mcp)))
    wb.selectedTab.foreach(workbook.setSelectedTab)
    evictFromCache(workbook)
    workbook
  }

  private def evictFromCache(wb: XSSFWorkbook) {
    cellStyleCache.remove(wb)
    dataFormatCache.remove(wb)
    fontCache.remove(wb)
  }

  //================= Cache processing ====================
  private def getCachedOrUpdate[K, V](cache: Cache[K, V], value: K, workbook: XSSFWorkbook)(newValue: => V): V = {
    val workbookCache = cache.getOrElseUpdate(workbook, collection.mutable.Map[K, V]())
    workbookCache.getOrElseUpdate(value, newValue)
  }

  implicit class XlsxBorderStyle(bs: CellBorderStyle) {
    def convertAsXlsx(): BorderStyle = convertBorderStyle(bs)
  }

  implicit class XlsxColor(c: Color) {
    def convertAsXlsx(): XSSFColor = convertColor(c)
  }

  implicit class XlsxCellFill(cf: CellFill) {
    def convertAsXlsx(): FillPatternType = convertCellFill(cf)
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
    def convertAsXlsx(): HorizontalAlignment = convertHorizontalAlignment(ha)
  }

  implicit class XlsxSheet(s: Sheet) {
    def convertAsXlsx(workbook: XSSFWorkbook): XSSFSheet = convertSheet(s, workbook)

    def convertAsXlsx(): XSSFWorkbook = Workbook(s).convertAsXlsx()

    def saveAsXlsx(fileName: String) {
      Workbook(s).saveAsXlsx(fileName)
    }
  }

  implicit class XlsxVerticalAlignment(va: CellVerticalAlignment) {
    def convertAsXlsx(): VerticalAlignment = convertVerticalAlignment(va)
  }

  implicit class XlsxWorkbook(workbook: Workbook) {
    def saveAsXlsx(fileName: String): Unit = {
      val stream = new FileOutputStream(fileName)
      try {
        val workbook = convertAsXlsx()
        workbook.write(stream)
      } finally {
        stream.flush()
        stream.close()
      }
    }

    def convertAsXlsx(): XSSFWorkbook = convertWorkbook(workbook)
  }

}
