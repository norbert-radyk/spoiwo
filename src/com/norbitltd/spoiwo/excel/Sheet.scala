package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}

object Sheet {

  def apply(rows: Row*) = apply(rows = rows.toList)

}

case class Sheet(name: String = "",
                 columns: List[Column] = Nil,
                 rows: List[Row] = Nil,
                 mergedRegions: List[CellRangeAddress] = Nil,
                 autoBreaks: Boolean = false,
                 printSetup: PrintSetup = PrintSetup.Default,
                 header: Header = Header.None,
                 footer: Footer = Footer.None) {

  def withSheetName(name: String) = copy(name = name)

  def withAutoBreaks(autoBreaks: Boolean) = copy(autoBreaks = autoBreaks)

  def withPrintSetup(printSetup : PrintSetup) = copy(printSetup = printSetup)

  def withHeader(header: Header) = copy(header = header)

  def withFooter(footer : Footer) = copy(footer = footer)



  def convert(workbook: XSSFWorkbook): XSSFSheet = {
    val sheetName = if (name.isEmpty) "Sheet " + (workbook.getNumberOfSheets + 1) else sheetName
    val sheet = workbook.createSheet(sheetName)
    initializeColumns(sheet)
    initializeRows(sheet)
    initializeMergedRegions(sheet)

    sheet.setAutobreaks(autoBreaks)
    sheet.set

    printSetup.applyTo(sheet)

    header.apply(sheet)
    footer.apply(sheet)
    sheet
  }

  def initializeRows(sheet: XSSFSheet) {
    rows.foreach(row => row.convert(sheet))
  }

  def initializeColumns(sheet: XSSFSheet) {
    columns.foreach(column => {
      val columnIndex = column.index
      sheet.setColumnWidth(columnIndex, 256 * column.width)
      sheet.setColumnHidden(columnIndex, column.hidden)
      sheet.setColumnGroupCollapsed(columnIndex, column.groupCollapsed)
    })
  }

  def initializeMergedRegions(sheet: XSSFSheet) {
    mergedRegions.foreach(mergedRegion => sheet.addMergedRegion(mergedRegion.convert()))
  }

  def save(fileName: String) {
    Workbook(this).save(fileName)
  }

}
