package com.norbitltd.spoiwo.model

object Table {

  def apply(
    cellRange: CellRange,
    id: java.lang.Long = null,
    name: String = null,
    displayName: String = null,
    columns: List[TableColumn] = Nil,
    style: TableStyle = null,
    enableAutoFilter: java.lang.Boolean = null
  ): Table =
    Table(
      cellRange        = cellRange,
      id               = Option(id).map(_.longValue),
      name             = Option(name),
      displayName      = Option(displayName),
      columns          = columns,
      style            = Option(style),
      enableAutoFilter = Option(enableAutoFilter).map(_.booleanValue)
    )
}

case class Table private(
  cellRange: CellRange,
  id: Option[Long],
  name: Option[String],
  displayName: Option[String],
  columns: List[TableColumn],
  style: Option[TableStyle],
  enableAutoFilter: Option[Boolean]
) {

  def withCellRange(cellRange: CellRange) =
    copy(cellRange = cellRange)

  def withId(id: Long) =
    copy(id = Some(id))

  def withName(name: String) =
    copy(name = Option(name))

  def withDisplayName(displayName: String) =
    copy(displayName = Option(displayName))

  def withColumns(columns: List[TableColumn]) =
    copy(columns = columns)

  def withStyle(style: TableStyle) =
    copy(style = Option(style))

  def withoutStyle =
    copy(style = None)

  def withAutoFilter =
    copy(enableAutoFilter = Some(true))

  def withoutAutoFilter =
    copy(enableAutoFilter = Some(false))
}