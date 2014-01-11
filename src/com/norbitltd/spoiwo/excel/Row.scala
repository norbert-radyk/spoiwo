package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.{XSSFRow, XSSFSheet}

object Row {

  val Empty = apply()

  def apply(cells: Cell*): Row = apply(cells = cells.toVector)

  def apply(cellStyle: CellStyle, cellValues: List[Any]): Row = {
    val cells = cellValues.map(value => value match {
      case stringValue: String => Cell(stringValue, cellStyle)
      case doubleValue: Double => Cell(doubleValue, cellStyle)
      case intValue: Int => Cell(intValue, cellStyle)
      case longValue: Long => Cell(longValue, cellStyle)
      case _ => throw new UnsupportedOperationException("Unable to construct cell from " + value.getClass + " type value!")
    })
    apply(cells = cells.toVector)
  }

}

case class Row(heightInPoints: Float = 10, cells: Vector[Cell]) {

  def withHeightInPoints(heightInPoints : Float) : Row = copy(heightInPoints = heightInPoints)

  def convert(sheet: XSSFSheet): XSSFRow = {
    val row = sheet.createRow(getNextRowNumber(sheet))
    cells.foreach(cell => cell.convert(row))
    row.setHeightInPoints(heightInPoints)
    row
  }

  def getNextRowNumber(sheet: XSSFSheet) = if (sheet.rowIterator().hasNext)
    sheet.getLastRowNum + 1
  else
    0

}



