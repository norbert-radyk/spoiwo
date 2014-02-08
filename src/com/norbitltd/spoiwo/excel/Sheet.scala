package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}

object Sheet extends Factory {

  private lazy val defaultName = ""
  private lazy val defaultColumns = Nil
  private lazy val defaultRows = Nil
  private lazy val default

  private lazy val name = defaultPOISheet

  val Blank = Sheet()

  def apply(rows: Row*): Sheet = apply(rows = rows.toList)

}

case class Sheet(name: String = "",
                 columns: List[Column] = Nil,
                 rows: List[Row] = Nil,
                 mergedRegions: List[CellRange] = Nil,
                 printSetup: PrintSetup = PrintSetup.Default,
                 header: Header = Header.None,
                 footer: Footer = Footer.None,
                 properties: SheetProperties = SheetProperties.Default,
                 margins: Margins = Margins.Default) {

  def withSheetName(name: String) = copy(name = name)

  def withColumns(columns: List[Column]): Sheet = copy(columns = columns)

  def withColumns(columns: Column*): Sheet = withColumns(columns.toList)

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

    printSetup.applyTo(sheet)
    header.applyTo(sheet)
    footer.applyTo(sheet)
    properties.applyTo(sheet)
    sheet
  }

  def initializeRows(sheet: XSSFSheet) {
    rows.foreach(row => row.convert(sheet))
  }

  def initializeColumns(sheet: XSSFSheet) {
    updateColumnsWithIndexes().foreach {
      _.applyTo(sheet)
    }
  }

  def initializeMergedRegions(sheet: XSSFSheet) {
    mergedRegions.foreach(mergedRegion => sheet.addMergedRegion(mergedRegion.convert()))
  }

  def save(fileName: String) {
    Workbook(this).save(fileName)
  }

  private def updateColumnsWithIndexes(): List[Column] = {
    val currentColumnIndexes = columns.map(_.index).flatten.toSet
    if (currentColumnIndexes.isEmpty) {
      columns.zipWithIndex.map {
        case (column, index) => column.withIndex(index)
      }
    } else if (currentColumnIndexes.size == columns.size) {
      columns
    } else {
      throw new IllegalArgumentException(
        "When explicitly specifying column index you are required to provide it " +
          "uniquely for all columns in this sheet definition!")
    }
  }

}
