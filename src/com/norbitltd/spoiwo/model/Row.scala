package com.norbitltd.spoiwo.model

import java.util.{Calendar, Date}

object Row extends Factory {

  private lazy val defaultCells = Nil
  private lazy val defaultHeight = defaultPOIRow.getHeight
  lazy val DefaultHeightInPoints = defaultPOIRow.getHeightInPoints
  private lazy val defaultIndex = -1
  private lazy val defaultStyle = CellStyle.Default
  private lazy val defaultZeroHeight = false

  val Empty = apply()

  def apply(cells: Iterable[Cell] = defaultCells,
            height: Short = defaultHeight,
            heightInPoints: Float = DefaultHeightInPoints,
            index: Int = defaultIndex,
            style: CellStyle = defaultStyle,
            zeroHeight: Boolean = defaultZeroHeight): Row =
    Row(cells = cells,
      height = wrap(height, defaultHeight),
      heightInPoints = wrap(heightInPoints, DefaultHeightInPoints),
      index = wrap(index, defaultIndex),
      style = wrap(style, defaultStyle),
      zeroHeight = wrap(zeroHeight, defaultZeroHeight)
    )

  def apply(cells: Cell*): Row =
    apply(cells = cells.toVector)

  def apply(index: Int, cell: Cell, cells: Cell*): Row =
    apply(index = index, cells = cell :: cells.toList)

  def apply(style: CellStyle, cell: Cell, cells: Cell*): Row =
    apply(style = style, cells = cell :: cells.toList)

  def apply(index: Int, style: CellStyle, cell: Cell, cells: Cell*): Row =
    apply(index = index, style = style, cells = cell :: cells.toList)

}

case class Row private(cells: Iterable[Cell],
                       height: Option[Short],
                       heightInPoints: Option[Float],
                       index: Option[Int],
                       style: Option[CellStyle],
                       zeroHeight: Option[Boolean]) {

  def withCells(cells: Cell*) =
    copy(cells = cells)

  def withCells(cells: Iterable[Cell]) =
    copy(cells = cells)

  def withCellValues(cellValues: List[Any]): Row = {
    val cells = cellValues.map {
      case stringValue: String => Cell(stringValue)
      case doubleValue: Double => Cell(doubleValue)
      case intValue: Int => Cell(intValue.toDouble)
      case longValue: Long => Cell(longValue.toDouble)
      case booleanValue: Boolean => Cell(booleanValue)
      case dateValue: Date => Cell(dateValue)
      case calendarValue: Calendar => Cell(calendarValue)
      case value => throw new UnsupportedOperationException("Unable to construct cell from " + value.getClass + " type value!")
    }
    copy(cells = cells.toVector)
  }

  def withCellValues(cellValues: Any*): Row = {
    val cells = cellValues.map {
      case stringValue: String => Cell(stringValue)
      case doubleValue: Double => Cell(doubleValue)
      case intValue: Int => Cell(intValue.toDouble)
      case longValue: Long => Cell(longValue.toDouble)
      case booleanValue: Boolean => Cell(booleanValue)
      case dateValue: Date => Cell(dateValue)
      case calendarValue: Calendar => Cell(calendarValue)
      case value => throw new UnsupportedOperationException("Unable to construct cell from " + value.getClass + " type value!")
    }
    copy(cells = cells.toVector)
  }

  def withHeight(height: Short) =
    copy(height = Option(height))

  def withHeightInPoints(heightInPoints: Float) =
    copy(heightInPoints = Option(heightInPoints))

  def withStyle(rowStyle: CellStyle) =
    copy(style = Option(rowStyle))

  def withZeroHeight(zeroHeight: Boolean) =
    copy(zeroHeight = Option(zeroHeight))
}



