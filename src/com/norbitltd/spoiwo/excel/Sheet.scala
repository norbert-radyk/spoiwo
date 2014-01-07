package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}

object Sheet {

}

case class Sheet(name: String, columns: List[Column] = Nil, rows: List[Row]) {

  def convert(workbook: XSSFWorkbook): XSSFSheet = {
    val sheet = workbook.createSheet(name)
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

}
