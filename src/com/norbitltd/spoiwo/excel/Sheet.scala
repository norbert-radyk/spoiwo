package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}

object Sheet {

  def apply(rows : Row*) = apply(rows = rows.toList)

}

case class Sheet(name: String = "", columns: List[Column] = Nil, rows: List[Row] = Nil) {

  def withSheetName(name : String) = copy(name = name)

  def convert(workbook: XSSFWorkbook): XSSFSheet = {
    val sheetName = if( name.isEmpty ) "Sheet " + (workbook.getNumberOfSheets + 1) else sheetName
    val sheet = workbook.createSheet(sheetName)
    initializeColumns(sheet)
    initializeRows(sheet)
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

  def save(fileName : String) {
    Workbook(this).save(fileName)
  }

}
