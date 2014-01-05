package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.ss.usermodel.Row

object XRow {

  val Empty = apply()

  def apply(cells : XCell*) : XRow = apply(cells.toVector)

  def apply(cellStyle : XCellStyle, cellValues : List[Any]) : XRow = {
    val cells = cellValues.map(value => value match {
      case stringValue : String => XCell(stringValue, cellStyle)
      case doubleValue : Double => XCell(doubleValue, cellStyle)
      case intValue : Int => XCell(intValue, cellStyle)
      case longValue : Long => XCell(longValue, cellStyle)
      case _ => throw new UnsupportedOperationException("Unable to construct cell from " + value.getClass + " type value!")
    })
    apply(cells.toVector)
  }

}

case class XRow(cells: Vector[XCell]) {

  def convert(sheet: XSSFSheet): Row = {
    val row = sheet.createRow(getNextRowNumber(sheet))
    cells.foreach(cell => cell.convert(row))
    row
  }

  def getNextRowNumber(sheet: XSSFSheet) = if (sheet.rowIterator().hasNext)
    sheet.getLastRowNum + 1
  else
    0

}



