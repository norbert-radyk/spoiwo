package com.norbitltd.spoiwo.natures.streaming.xlsx

import java.io.{FileOutputStream, OutputStream}

import com.norbitltd.spoiwo.model.enums._
import com.norbitltd.spoiwo.model.{BooleanCell, CalendarCell, DateCell, FormulaCell, NumericCell, StringCell, _}
import com.norbitltd.spoiwo.natures.xlsx.BaseXlsx
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxEnumConversions._
import org.apache.poi.ss.usermodel
import org.apache.poi.ss.usermodel.{BorderStyle, CellType, FillPatternType, HorizontalAlignment, ReadingOrder, VerticalAlignment}
import org.apache.poi.xssf.streaming.{SXSSFCell, SXSSFRow, SXSSFSheet, SXSSFWorkbook}
import org.apache.poi.xssf.usermodel._

object Model2XlsxConversions extends BaseXlsx {

  private type Cache[K, V] = collection.mutable.Map[SXSSFWorkbook, collection.mutable.Map[K, V]]
  private lazy val cellStyleCache = Cache[CellStyle, XSSFCellStyle]()
  private lazy val dataFormatCache = collection.mutable.Map[SXSSFWorkbook, XSSFDataFormat]()
  private lazy val fontCache = Cache[Font, XSSFFont]()

  private def Cache[K, V]() = collection.mutable.Map[SXSSFWorkbook, collection.mutable.Map[K, V]]()


  private[xlsx] def convertCell(modelSheet: Sheet,
                                modelColumns: Map[Int, Column],
                                modelRow: Row,
                                c: Cell,
                                row: SXSSFRow): SXSSFCell = {
    val cellNumber = c.index.getOrElse(if (row.getLastCellNum < 0) 0 else row.getLastCellNum.toInt)
    val cell = Option(row.getCell(cellNumber)).getOrElse(row.createCell(cellNumber))
    if (cell.getCellType == CellType.FORMULA) {
      cell.setCellFormula(null)
    }

    val cellWithStyle = mergeStyle(c, modelRow.style, modelColumns.get(cellNumber).flatMap(_.style), modelSheet.style)
    cellWithStyle.style.foreach(s =>
      cell.setCellStyle(convertCellStyle(s, cell.getRow.getSheet.getWorkbook.asInstanceOf[SXSSFWorkbook])))

    c match {
      case BlankCell(_, _, _)               => cell.setCellValue(null: String)
      case StringCell(value, _, _, _)       => cell.setCellValue(value)
      case FormulaCell(formula, _, _, _)    => cell.setCellFormula(formula)
      case NumericCell(value, _, _, _)      => cell.setCellValue(value)
      case BooleanCell(value, _, _, _)      => cell.setCellValue(value)
      case DateCell(value, _, _, _)         => setDateCell(c, cell, value)
      case CalendarCell(value, _, _, _)     => setCalendarCell(c, cell, value)
      case HyperLinkUrlCell(value, _, _, _) => setHyperLinkUrlCell(cell, value, row)
    }
    cell
  }

  private def convertCellDataFormat(cdf: CellDataFormat, workbook: SXSSFWorkbook, cellStyle: usermodel.CellStyle): Unit = {
    cdf.formatString.foreach(formatString => {
      val format = dataFormatCache.getOrElseUpdate(workbook, workbook.getXSSFWorkbook.createDataFormat())
      val formatIndex = format.getFormat(formatString)
      cellStyle.setDataFormat(formatIndex)
    })
  }

  private[xlsx] def convertCellStyle(cs: CellStyle, workbook: SXSSFWorkbook): XSSFCellStyle =
    getCachedOrUpdate(cellStyleCache, cs, workbook) {
      val cellStyle = workbook.getXSSFWorkbook.createCellStyle()
      cs.borders.foreach(b => convertCellBorders(b, cellStyle))
      cs.dataFormat.foreach(df => convertCellDataFormat(df, workbook, cellStyle))
      cs.font.foreach(f => cellStyle.setFont(convertFont(f, workbook)))
      cs.fillPattern.foreach(fp => cellStyle.setFillPattern(convertCellFill(fp)))
      cs.fillBackgroundColor.foreach(c => cellStyle.setFillBackgroundColor(convertColor(c)))
      cs.fillForegroundColor.foreach(c => cellStyle.setFillForegroundColor(convertColor(c)))

      cs.readingOrder.foreach(ro => cellStyle.setReadingOrder(convertReadingOrder(ro)))
      cs.horizontalAlignment.foreach(ha => cellStyle.setAlignment(convertHorizontalAlignment(ha)))
      cs.verticalAlignment.foreach(va => cellStyle.setVerticalAlignment(convertVerticalAlignment(va)))

      cs.hidden.foreach(cellStyle.setHidden)
      cs.indention.foreach(cellStyle.setIndention)
      cs.locked.foreach(cellStyle.setLocked)
      cs.rotation.foreach(cellStyle.setRotation)
      cs.wrapText.foreach(cellStyle.setWrapText)
      cellStyle
    }

  private def convertFooter(f: Footer, sheet: SXSSFSheet): Unit = {
    f.left.foreach(sheet.getFooter.setLeft)
    f.center.foreach(sheet.getFooter.setCenter)
    f.right.foreach(sheet.getFooter.setRight)
  }

  private[xlsx] def convertFont(f: Font, workbook: SXSSFWorkbook): XSSFFont =
    getCachedOrUpdate(fontCache, f, workbook) {
      val font = workbook.getXSSFWorkbook.createFont()
      convertFont(f, font)
    }

  private def convertHeader(h: Header, sheet: SXSSFSheet): Unit = {
    h.left.foreach(sheet.getHeader.setLeft)
    h.center.foreach(sheet.getHeader.setCenter)
    h.right.foreach(sheet.getHeader.setRight)
  }

  private[xlsx] def convertRow(modelColumns: Map[Int, Column],
                               modelRow: Row,
                               modelSheet: Sheet,
                               sheet: SXSSFSheet): SXSSFRow = {
    validateCells(modelRow)
    val indexNumber = modelRow.index.getOrElse(if (sheet.rowIterator().hasNext) sheet.getLastRowNum + 1 else 0)
    val row = Option(sheet.getRow(indexNumber)).getOrElse(sheet.createRow(indexNumber))
    modelRow.height.foreach(h => row.setHeightInPoints(h.toPoints))
    modelRow.style.foreach(s => row.setRowStyle(convertCellStyle(s, row.getSheet.getWorkbook)))
    modelRow.hidden.foreach(row.setZeroHeight)
    modelRow.cells.foreach(cell => convertCell(modelSheet, modelColumns, modelRow, cell, row))
    row
  }

  private[xlsx] def convertSheet(s: Sheet, workbook: SXSSFWorkbook): SXSSFSheet = {
    s.validate()
    writeToExistingSheet(s, workbook.createSheet(s.nameIn(workbook)))
  }

  private[xlsx] def writeToExistingSheet(s: Sheet, sheet: SXSSFSheet): SXSSFSheet = {
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

    sheet
  }

  override def setTabColor(sheet: usermodel.Sheet, color: XSSFColor): Unit = {
    sheet.asInstanceOf[SXSSFSheet].setTabColor(color)
  }

  override def additionalPrintSetup(printSetup: PrintSetup, sheetPs: usermodel.PrintSetup): Unit = {
    if (printSetup != PrintSetup.Default) {
      val ps = sheetPs.asInstanceOf[XSSFPrintSetup]
      printSetup.pageOrder.foreach(po => ps.setPageOrder(convertPageOrder(po)))
      printSetup.paperSize.foreach(px => ps.setPaperSize(convertPaperSize(px)))
    }
  }

  private[xlsx] def validateTables(modelSheet: Sheet): Unit = {
    if (modelSheet.tables.nonEmpty) {
      throw new IllegalStateException("createTable is not supported by SXSSF right now")
    }
  }

  private[xlsx] def convertWorkbook(wb: Workbook): SXSSFWorkbook = {
    writeToExistingWorkbook(wb, new SXSSFWorkbook())
  }

  private[xlsx] def writeToExistingWorkbook(wb: Workbook, workbook: SXSSFWorkbook): SXSSFWorkbook = {
    wb.sheets.foreach { s =>
      s.validate()
      val sheetName = s.nameIn(workbook)
      writeToExistingSheet(s, Option(workbook.getSheet(sheetName)).getOrElse(workbook.createSheet(sheetName)))
    }

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
  private def evictFromCache(wb: SXSSFWorkbook): Unit = {
    cellStyleCache.remove(wb)
    dataFormatCache.remove(wb)
    fontCache.remove(wb)
  }

  //================= Cache processing ====================
  private def getCachedOrUpdate[K, V](cache: Cache[K, V], value: K, workbook: SXSSFWorkbook)(newValue: => V): V = {
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
    def convertAsXlsx(cell: SXSSFCell): XSSFCellStyle = convertAsXlsx(cell.getRow.asInstanceOf[SXSSFRow])

    def convertAsXlsx(row: SXSSFRow): XSSFCellStyle = convertAsXlsx(row.getSheet)

    def convertAsXlsx(sheet: SXSSFSheet): XSSFCellStyle = convertAsXlsx(sheet.getWorkbook)

    def convertAsXlsx(workbook: SXSSFWorkbook): XSSFCellStyle = convertCellStyle(cs, workbook)
  }

  implicit class XlsxFont(f: Font) {
    def convertAsXlsx(cell: SXSSFCell): XSSFFont = convertAsXlsx(cell.getRow.asInstanceOf[SXSSFRow])

    def convertAsXlsx(row: SXSSFRow): XSSFFont = convertAsXlsx(row.getSheet)

    def convertAsXlsx(sheet: SXSSFSheet): XSSFFont = convertAsXlsx(sheet.getWorkbook)

    def convertAsXlsx(workbook: SXSSFWorkbook): XSSFFont = convertFont(f, workbook)
  }

  implicit class XlsxHorizontalAlignment(ha: CellHorizontalAlignment) {
    def convertAsXlsx(): HorizontalAlignment = convertHorizontalAlignment(ha)
  }

  sealed trait XlsxExport {
    def saveAsXlsx(fileName: String): Unit

    def writeToOutputStream[T <: OutputStream](stream: T): T
  }

  implicit class XlsxReadingOrder(ro: CellReadingOrder) {
    def convertAsXlsx(): ReadingOrder = convertReadingOrder(ro)
  }

  implicit class XlsxSheet(s: Sheet) extends XlsxExport {
    def validate(): Unit = {
      validateRows(s)
      validateTables(s)
    }

    def nameIn(workbook: SXSSFWorkbook): String = s.name.getOrElse("Sheet" + (workbook.getNumberOfSheets + 1))

    def writeToExisting(existingSheet: SXSSFSheet): Unit = writeToExistingSheet(s, existingSheet)

    def convertAsXlsx(workbook: SXSSFWorkbook): SXSSFSheet = convertSheet(s, workbook)

    def convertAsXlsx(): SXSSFWorkbook = Workbook(s).convertAsXlsx()

    override def saveAsXlsx(fileName: String): Unit = {
      Workbook(s).saveAsXlsx(fileName)
    }

    override def writeToOutputStream[T <: OutputStream](stream: T): T = Workbook(s).writeToOutputStream(stream)
  }

  implicit class XlsxVerticalAlignment(va: CellVerticalAlignment) {
    def convertAsXlsx(): VerticalAlignment = convertVerticalAlignment(va)
  }

  implicit class XlsxWorkbook(workbook: Workbook) extends XlsxExport {
    def writeToExisting(existingWorkBook: SXSSFWorkbook): Unit = writeToExistingWorkbook(workbook, existingWorkBook)

    override def saveAsXlsx(fileName: String): Unit = writeToOutputStream(new FileOutputStream(fileName))

    override def writeToOutputStream[T <: OutputStream](stream: T): T =
      try {
        convertAsXlsx().write(stream)
        stream
      } finally {
        stream.flush()
        stream.close()
      }

    def convertAsXlsx(): SXSSFWorkbook = convertWorkbook(workbook)
  }

}
