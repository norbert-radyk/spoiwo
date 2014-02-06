package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.{XSSFRow, XSSFSheet}

object Row extends Factory {

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

case class Row(heightInPoints: Float = 10,
               rowStyle: CellStyle = CellStyle.Default,
               firstCellColumn: Int = 0,
               cells: Iterable[Cell] = Nil) {

  val isZeroHeightRow = (heightInPoints == 0)

  def withHeightInPoints(heightInPoints: Float) =
    copy(heightInPoints = heightInPoints)

  def withRowStyle(rowStyle: CellStyle) =
    copy(rowStyle = rowStyle)

  def withFirstCellColumn(firstCellColumn: Int) =
    copy(firstCellColumn = firstCellColumn)

  def withCells(cells : Cell*) =
    copy(cells = cells)

  def withCells(cells : Iterable[Cell]) =
    copy(cells = cells)

  def convert(sheet: XSSFSheet): XSSFRow = {
    val row = sheet.createRow(getNextRowNumber(sheet))
    createInitialEmptyCellsFor(row)
    cells.foreach(cell => cell.convert(row))
    row.setHeightInPoints(heightInPoints)
    row.setRowStyle(rowStyle.convert(row))
    row.setZeroHeight(isZeroHeightRow)
    row
  }

  private def createInitialEmptyCellsFor(row: XSSFRow) {
    (0 to firstCellColumn).foreach(columnNumber => Cell.Empty.convert(row))
  }

  private def getNextRowNumber(sheet: XSSFSheet) = if (sheet.rowIterator().hasNext)
    sheet.getLastRowNum + 1
  else
    0

}



