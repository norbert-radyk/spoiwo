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
      cellRange = cellRange,
      id = Option(id).map(_.longValue),
      name = Option(name),
      displayName = Option(displayName),
      columns = columns,
      style = Option(style),
      enableAutoFilter = Option(enableAutoFilter).map(_.booleanValue)
    )
}

case class Table private (
    cellRange: CellRange,
    id: Option[Long],
    name: Option[String],
    displayName: Option[String],
    columns: List[TableColumn],
    style: Option[TableStyle],
    enableAutoFilter: Option[Boolean]
) {

  def withCellRange(cellRange: CellRange): Table =
    copy(cellRange = cellRange)

  def withId(id: Long): Table =
    copy(id = Some(id))

  def withName(name: String): Table =
    copy(name = Option(name))

  def withDisplayName(displayName: String): Table =
    copy(displayName = Option(displayName))

  def withColumns(columns: List[TableColumn]): Table =
    copy(columns = columns)

  def withStyle(style: TableStyle): Table =
    copy(style = Option(style))

  def withoutStyle: Table =
    copy(style = None)

  def withAutoFilter: Table =
    copy(enableAutoFilter = Some(true))

  def withoutAutoFilter: Table =
    copy(enableAutoFilter = Some(false))
}
