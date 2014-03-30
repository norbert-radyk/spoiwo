package com.norbitltd.spoiwo.model

import java.util.{Calendar, Date}
import org.joda.time.{DateTime, LocalDate}

object Row {

  val Empty = apply()

  def apply(cells: Iterable[Cell] = Nil,
            height: Height = null,
            index: java.lang.Integer = null,
            style: CellStyle = null,
            hidden: java.lang.Boolean = null): Row =
    Row(cells = cells,
      height = Option(height),
      index = Option(index).map(_.intValue),
      style = Option(style),
      hidden = Option(hidden).map(_.booleanValue)
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
                       height: Option[Height],
                       index: Option[Int],
                       style: Option[CellStyle],
                       hidden: Option[Boolean]) {

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
      case dateValue: LocalDate => Cell(dateValue)
      case dateValue: DateTime => Cell(dateValue)
      case calendarValue: Calendar => Cell(calendarValue)
      case value => throw new UnsupportedOperationException("Unable to construct cell from " + value.getClass + " type value!")
    }
    copy(cells = cells.toVector)
  }

  def withHeight(height: Height) =
    copy(height = Option(height))

  def withoutHeight =
    copy(height = None)

  def withStyle(rowStyle: CellStyle) =
    copy(style = Option(rowStyle))

  def withoutStyle =
    copy(style = None)

  def withHidden =
    copy(hidden = Some(true))

  def withoutHidden =
    copy(hidden = Some(false))
}



