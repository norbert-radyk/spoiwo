package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}

object Sheet extends Factory {

  private lazy val name = defaultPOISheet

  val Blank = Sheet()

  def apply(rows: Row*): Sheet = apply(rows = rows.toList)

}

case class Sheet(name: String = "",
                 columns: Map[Short, Column] = Map(),
                 rows: List[Row] = Nil,
                 mergedRegions: List[CellRangeAddress] = Nil,
                 printSetup: PrintSetup = PrintSetup.Default,
                 header: Header = Header.None,
                 footer: Footer = Footer.None,
                 properties: SheetProperties = SheetProperties.Default,
                  margins : Margins = Margins.Default) {

  def withSheetName(name: String) = copy(name = name)

  def withColumns(columns: Map[Short, Column]): Sheet = copy(columns = columns)

  def withColumns(columns: Iterable[Column]): Sheet = withColumns(
    columns.zipWithIndex.map {
      case (column, index) => index.toShort -> column
    }.toMap
  )

  def withColumns(columns: Column*): Sheet = withColumns(columns)

  def withRows(rows: Iterable[Row]): Sheet = copy(rows = rows.toList)

  def withRows(rows: Row*): Sheet = withRows(rows)

  def withPrintSetup(printSetup: PrintSetup) = copy(printSetup = printSetup)

  def withHeader(header: Header) = copy(header = header)

  def withFooter(footer: Footer) = copy(footer = footer)

  def convert(workbook: XSSFWorkbook): XSSFSheet = {
    val sheetName = if (name.isEmpty) "Sheet " + (workbook.getNumberOfSheets + 1) else name
    val sheet = workbook.createSheet(sheetName)
    initializeColumns(sheet)
    initializeRows(sheet)
    initializeMergedRegions(sheet)

    //TODO Add sheet properties
    //TODO on the sheet level: workbook.setPrintArea()
    printSetup.applyTo(sheet)

    header.apply(sheet)
    footer.apply(sheet)
    properties.applyTo(sheet)
    sheet
  }

  def initializeRows(sheet: XSSFSheet) {
    rows.foreach(row => row.convert(sheet))
  }

  def initializeColumns(sheet: XSSFSheet) {
    columns.foreach {
      case (index, column) => column.applyTo(index, sheet)
    }
  }

  def initializeMergedRegions(sheet: XSSFSheet) {
    mergedRegions.foreach(mergedRegion => sheet.addMergedRegion(mergedRegion.convert()))
  }

  def save(fileName: String) {
    Workbook(this).save(fileName)
  }

}
