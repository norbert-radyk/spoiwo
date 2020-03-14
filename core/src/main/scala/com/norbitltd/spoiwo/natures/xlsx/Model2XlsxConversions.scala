package com.norbitltd.spoiwo.natures.xlsx

import java.io.{FileInputStream, FileOutputStream, OutputStream}
import com.norbitltd.spoiwo.model.enums._
import com.norbitltd.spoiwo.model.{BooleanCell, CalendarCell, DateCell, FormulaCell, NumericCell, StringCell, _}
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxEnumConversions._
import org.apache.poi.ss.usermodel
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType
import org.apache.poi.ss.usermodel.{
  BorderStyle,
  CellType,
  FillPatternType,
  ReadingOrder,
  HorizontalAlignment,
  VerticalAlignment
}
import org.apache.poi.util.IOUtils
import org.apache.poi.xssf.usermodel._
import org.openxmlformats.schemas.spreadsheetml.x2006.main.{CTTable, CTTableColumns, CTTableStyleInfo}

object Model2XlsxConversions extends BaseXlsx {

  private type Cache[K, V] = collection.mutable.Map[XSSFWorkbook, collection.mutable.Map[K, V]]
  private lazy val cellStyleCache = Cache[CellStyle, XSSFCellStyle]()
  private lazy val dataFormatCache = collection.mutable.Map[XSSFWorkbook, XSSFDataFormat]()
  private lazy val fontCache = Cache[Font, XSSFFont]()

  private def Cache[K, V]() = collection.mutable.Map[XSSFWorkbook, collection.mutable.Map[K, V]]()

  private[xlsx] def convertCell(
      modelSheet: Sheet,
      modelColumns: Map[Int, Column],
      modelRow: Row,
      c: Cell,
      row: XSSFRow
  ): XSSFCell = {
    val cellNumber = c.index.getOrElse(if (row.getLastCellNum < 0) 0 else row.getLastCellNum.toInt)
    val cell = Option(row.getCell(cellNumber)).getOrElse(row.createCell(cellNumber))
    if (cell.getCellType == CellType.FORMULA) {
      cell.setCellFormula(null)
    }

    val cellWithStyle = mergeStyle(c, modelRow.style, modelColumns.get(cellNumber).flatMap(_.style), modelSheet.style)
    cellWithStyle.style.foreach(s => cell.setCellStyle(convertCellStyle(s, cell.getRow.getSheet.getWorkbook)))

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

  private def convertCellDataFormat(cdf: CellDataFormat, workbook: XSSFWorkbook, cellStyle: XSSFCellStyle): Unit = {
    cdf.formatString.foreach(formatString => {
      val format = dataFormatCache.getOrElseUpdate(workbook, workbook.createDataFormat())
      val formatIndex = format.getFormat(formatString)
      cellStyle.setDataFormat(formatIndex)
    })
  }

  private[xlsx] def convertCellStyle(cs: CellStyle, workbook: XSSFWorkbook): XSSFCellStyle =
    getCachedOrUpdate(cellStyleCache, cs, workbook) {
      val cellStyle = workbook.createCellStyle()
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

  private def convertFooter(f: Footer, sheet: XSSFSheet): Unit = {
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
      convertFont(f, font)
    }

  private def convertHeader(h: Header, sheet: XSSFSheet): Unit = {
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

  private[xlsx] def convertRow(
      modelColumns: Map[Int, Column],
      modelRow: Row,
      modelSheet: Sheet,
      sheet: XSSFSheet
  ): XSSFRow = {
    validateCells(modelRow)
    val indexNumber = modelRow.index.getOrElse(if (sheet.rowIterator().hasNext) sheet.getLastRowNum + 1 else 0)
    val row = Option(sheet.getRow(indexNumber)).getOrElse(sheet.createRow(indexNumber))
    modelRow.height.foreach(h => row.setHeightInPoints(h.toPoints))
    modelRow.style.foreach(s => row.setRowStyle(convertCellStyle(s, row.getSheet.getWorkbook)))
    modelRow.hidden.foreach(row.setZeroHeight)
    modelRow.cells.foreach(cell => convertCell(modelSheet, modelColumns, modelRow, cell, row))
    row
  }

  private[xlsx] def convertSheet(s: Sheet, workbook: XSSFWorkbook): XSSFSheet = {
    s.validate()
    writeToExistingSheet(s, workbook.createSheet(s.nameIn(workbook)))
  }

  private[xlsx] def writeToExistingSheet(s: Sheet, sheet: XSSFSheet): XSSFSheet = {
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
    tables.foreach(tbl => convertTable(tbl, sheet))
    s.images.foreach(img => addImage(img, sheet))

    sheet
  }

  override def setTabColor(sheet: usermodel.Sheet, color: XSSFColor): Unit = {
    sheet.asInstanceOf[XSSFSheet].setTabColor(color)
  }

  override def additionalPrintSetup(printSetup: PrintSetup, sheetPs: usermodel.PrintSetup): Unit = {
    if (printSetup != PrintSetup.Default) {
      val ps = sheetPs.asInstanceOf[XSSFPrintSetup]
      printSetup.pageOrder.foreach(po => ps.setPageOrder(convertPageOrder(po)))
      printSetup.paperSize.foreach(px => ps.setPaperSize(convertPaperSize(px)))
    }
  }

  private[xlsx] def validateTables(modelSheet: Sheet): Unit = {
    val ids = modelSheet.tables.flatMap(_.id)
    if (ids.size != ids.toSet.size)
      throw new IllegalArgumentException("Specified table ids need to be unique.")
  }

  private[xlsx] def updateTablesWithIds(modelSheet: Sheet): List[Table] = {
    modelSheet.tables.zipWithIndex.map {
      case (table, index) => if (table.id.isDefined) table else table.withId(index + 1)
    }
  }

  private[xlsx] def convertTable(modelTable: Table, sheet: XSSFSheet): XSSFTable = {
    validateTableColumns(modelTable)

    val tableId = modelTable.id.getOrElse {
      throw new IllegalArgumentException(
        "Undefined table id! " +
          "Something went terribly wrong as it should have been derived if not specified explicitly!"
      )
    }

    val displayName = modelTable.displayName.getOrElse(s"Table$tableId")
    val name = modelTable.name.getOrElse(s"ct_table_$tableId")

    val table = sheet.createTable()
    table.setDisplayName(displayName)
    table.setName(name)
    val ctTable = table.getCTTable
    ctTable.setId(tableId)
    setTableReference(modelTable, ctTable)
    convertTableColumns(modelTable, ctTable)
    modelTable.style.foreach(convertTableStyle(_, ctTable))
    modelTable.enableAutoFilter.foreach(af => if (af) ctTable.addNewAutoFilter())
    table
  }

  private[xlsx] def validateTableColumns(modelTable: Table): Unit = {

    def insufficientColumnsDefined = {
      val (sCol, eCol) = modelTable.cellRange.columnRange
      val neededColumns = (eCol - sCol) + 1
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
      val (sCol, eCol) = modelTable.cellRange.columnRange
      val neededColumns = (eCol - sCol) + 1
      (0 until neededColumns) map { index =>
        val columnId = index + 1
        TableColumn(
          name = s"TableColumn$columnId",
          id = columnId.toLong
        )
      }
    }

    val modelColumns = if (modelTable.columns.nonEmpty) modelTable.columns else generateColumns
    val columns = ctTable.addNewTableColumns()
    columns.setCount(modelColumns.size)
    modelColumns.foreach { mc =>
      val column = columns.addNewTableColumn()
      column.setName(mc.name)
      column.setId(mc.id)
    }
    columns
  }

  private[xlsx] def convertImages(images: List[Image], sheet: XSSFSheet) =
    images.foldLeft(sheet)((sheet, image) => addImage(image, sheet))

  private[xlsx] def addImage(image: Image, sheet: XSSFSheet) = {
    val wb = sheet.getWorkbook

    val imageFormat =
      if (image.filePath.toLowerCase.endsWith(".jpg"))
        org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_JPEG
      else if (image.filePath.toLowerCase.endsWith(".png"))
        org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_PNG
      else throw new IllegalArgumentException(s"File format is not supported")

    val is = new FileInputStream(image.filePath)
    val bytes = IOUtils.toByteArray(is)
    val pictureIdx = wb.addPicture(bytes, imageFormat)
    val helper = wb.getCreationHelper
    val drawing = sheet.createDrawingPatriarch

    val anchor = helper.createClientAnchor
    anchor.setAnchorType(AnchorType.DONT_MOVE_AND_RESIZE)

    val (fromColumn, toColumn) = image.region.columnRange
    val (fromRow, toRow) = image.region.rowRange

    anchor.setCol1(fromColumn)
    anchor.setCol2(toColumn)
    anchor.setRow1(fromRow)
    anchor.setRow2(toRow)

    val pict = drawing.createPicture(anchor, pictureIdx)
    pict.resize()
    sheet
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
    writeToExistingWorkbook(wb, new XSSFWorkbook())
  }

  private[xlsx] def writeToExistingWorkbook(wb: Workbook, workbook: XSSFWorkbook): XSSFWorkbook = {
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
  private def evictFromCache(wb: XSSFWorkbook): Unit = {
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

    def nameIn(workbook: XSSFWorkbook): String = s.name.getOrElse("Sheet" + (workbook.getNumberOfSheets + 1))

    def writeToExisting(existingSheet: XSSFSheet): Unit = writeToExistingSheet(s, existingSheet)

    def convertAsXlsx(workbook: XSSFWorkbook): XSSFSheet = convertSheet(s, workbook)

    def convertAsXlsx(): XSSFWorkbook = Workbook(s).convertAsXlsx()

    override def saveAsXlsx(fileName: String): Unit = {
      Workbook(s).saveAsXlsx(fileName)
    }

    override def writeToOutputStream[T <: OutputStream](stream: T): T = Workbook(s).writeToOutputStream(stream)
  }

  implicit class XlsxVerticalAlignment(va: CellVerticalAlignment) {
    def convertAsXlsx(): VerticalAlignment = convertVerticalAlignment(va)
  }

  implicit class XlsxWorkbook(workbook: Workbook) extends XlsxExport {
    def writeToExisting(existingWorkBook: XSSFWorkbook): Unit = writeToExistingWorkbook(workbook, existingWorkBook)

    override def saveAsXlsx(fileName: String): Unit = writeToOutputStream(new FileOutputStream(fileName))

    override def writeToOutputStream[T <: OutputStream](stream: T): T =
      try {
        convertAsXlsx().write(stream)
        stream
      } finally {
        stream.flush()
        stream.close()
      }

    def convertAsXlsx(): XSSFWorkbook = convertWorkbook(workbook)
  }

}
