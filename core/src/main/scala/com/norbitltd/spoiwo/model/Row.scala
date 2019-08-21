package com.norbitltd.spoiwo.model

import java.util.{Calendar, Date}
import java.time.{LocalDate => JLocalDate, LocalDateTime => JLocalDateTime}
import org.joda.time.{DateTime, LocalDate}

object Row {

  val Empty: Row = apply()

  def apply(cells: Iterable[Cell] = Nil,
            height: Height = null,
            index: java.lang.Integer = null,
            style: CellStyle = null,
            hidden: java.lang.Boolean = null): Row =
    Row(cells = cells,
        height = Option(height),
        index = Option(index).map(_.intValue),
        style = Option(style),
        hidden = Option(hidden).map(_.booleanValue))

  def apply(cells: Cell*): Row =
    apply(cells = cells.toVector)

  def apply(index: Int, cell: Cell, cells: Cell*): Row =
    apply(index = index, cells = cell :: cells.toList)

  def apply(style: CellStyle, cell: Cell, cells: Cell*): Row =
    apply(style = style, cells = cell :: cells.toList)

  def apply(index: Int, style: CellStyle, cell: Cell, cells: Cell*): Row =
    apply(index = index, style = style, cells = cell :: cells.toList)

}

case class Row private (cells: Iterable[Cell],
                        height: Option[Height],
                        index: Option[Int],
                        style: Option[CellStyle],
                        hidden: Option[Boolean]) {

  override def toString: String =
    "Row(" + List(
      if (cells.isEmpty) None else Option("cells=" + cells.map(_.toString).mkString(", ")),
      height.map("height=" + _),
      index.map("index=" + _),
      style.map("style=" + _),
      hidden.map("hidden=" + _)
    ).flatten.mkString(", ") + ")"

  def withCells(cells: Cell*): Row =
    copy(cells = cells)

  def withCells(cells: Iterable[Cell]): Row =
    copy(cells = cells)

  def withCellValues(cellValues: List[Any]): Row = {
    val cells = cellValues.map {
      case stringValue: String           => Cell(stringValue)
      case doubleValue: Double           => Cell(doubleValue)
      case decimalValue: BigDecimal      => Cell(decimalValue)
      case intValue: Int                 => Cell(intValue)
      case longValue: Long               => Cell(longValue)
      case booleanValue: Boolean         => Cell(booleanValue)
      case dateValue: Date               => Cell(dateValue)
      case dateValue: LocalDate          => Cell(dateValue)
      case dateValue: DateTime           => Cell(dateValue)
      case dateValue: JLocalDate         => Cell(dateValue)
      case dateTimeValue: JLocalDateTime => Cell(dateTimeValue)
      case calendarValue: Calendar       => Cell(calendarValue)
      case hyperLinkUrl: HyperLinkUrl    => Cell(hyperLinkUrl)
      case value =>
        throw new UnsupportedOperationException("Unable to construct cell from " + value.getClass + " type value!")
    }
    copy(cells = cells.toVector)
  }

  def withCellValues(cellValues: Any*): Row =
    withCellValues(cellValues.toList)

  def addCell(cell: Cell): Row =
    copy(cells = cells ++ List(cell))

  def addCells(additionalCells: Iterable[Cell]): Row =
    copy(cells = cells ++ additionalCells)

  def removeCell(cell: Cell): Row =
    copy(cells = cells.filter(_ != cell))

  def removeCells(whereCondition: Cell => Boolean): Row =
    copy(cells = cells.filter(whereCondition))

  def withHeight(height: Height): Row =
    copy(height = Option(height))

  def withoutHeight: Row =
    copy(height = None)

  def withStyle(rowStyle: CellStyle): Row =
    copy(style = Option(rowStyle))

  def withoutStyle: Row =
    copy(style = None)

  def withHidden: Row =
    copy(hidden = Some(true))

  def withoutHidden: Row =
    copy(hidden = Some(false))
}
