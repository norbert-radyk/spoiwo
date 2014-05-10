package com.norbitltd.spoiwo.model

object Sheet {

  val Blank = Sheet()

  def apply(name: String = null,
            columns: List[Column] = Nil,
            rows: List[Row] = Nil,
            mergedRegions: List[CellRange] = Nil,
            printSetup: PrintSetup = null,
            header: Header = null,
            footer: Footer = null,
            properties: SheetProperties = null,
            margins: Margins = null,
            paneAction: PaneAction = null,
            repeatingRows: RowRange = null,
            repeatingColumns: ColumnRange = null,
            style: CellStyle = null): Sheet =
    Sheet(
      name = Option(name),
      columns = columns,
      rows = rows,
      mergedRegions = mergedRegions,
      printSetup = Option(printSetup),
      header = Option(header),
      footer = Option(footer),
      properties = Option(properties),
      margins = Option(margins),
      paneAction = Option(paneAction),
      repeatingRows = Option(repeatingRows),
      repeatingColumns = Option(repeatingColumns),
      style = Option(style)
    )

  def apply(rows: Row*): Sheet = apply(rows = rows.toList)

  def apply(name: String, row: Row, rows: Row*): Sheet = apply(name = name, rows = row :: rows.toList)

}

case class Sheet private(
                          name: Option[String],
                          columns: List[Column],
                          rows: List[Row],
                          mergedRegions: List[CellRange],
                          printSetup: Option[PrintSetup],
                          header: Option[Header],
                          footer: Option[Footer],
                          properties: Option[SheetProperties],
                          margins: Option[Margins],
                          paneAction: Option[PaneAction],
                          repeatingRows: Option[RowRange],
                          repeatingColumns: Option[ColumnRange],
                          style : Option[CellStyle]) {

  override def toString = "Sheet(" + List(
    name.map("name=" + _),
    if (columns.isEmpty) None else Option("columns=\n\t" + columns.map(_.toString).mkString("\n\t")),
    if (rows.isEmpty) None else Option("rows=\n\t" + rows.map(_.toString).mkString("\n\t")),
    if (mergedRegions.isEmpty) None else Option("merged regions=\n\t" + mergedRegions.map(_.toString).mkString("\n\t")),
    printSetup.map("print setup=" + _),
    header.map("header=" + _),
    footer.map("footer=" + _),
    properties.map("properties=" + _),
    margins.map("margins=" + _),
    paneAction.map("pane action=" + _),
    repeatingRows.map("repeating rows=" + _),
    repeatingColumns.map("repeating columns=" + _),
    style.map("style=" + _)
  ).flatten.mkString(",\n") + ")"

  def withSheetName(name: String) =
    copy(name = Option(name))

  def withColumns(columns: List[Column]): Sheet =
    copy(columns = columns)

  def withColumns(columns: Column*): Sheet =
    withColumns(columns.toList)

  def addColumn(column : Column) =
    copy(columns = columns ++ List(column))

  def addColumns(additionalColumns : Iterable[Column]) =
    copy(columns = columns ++ additionalColumns)

  def removeColumn(column : Column) =
    copy(columns = columns.filter(_ != column))

  def removeColumns(whereCondition : Column => Boolean) =
    copy(columns = columns.filter(whereCondition))

  def withRows(rows: Iterable[Row]): Sheet =
    copy(rows = rows.toList)

  def withRows(rows: Row*): Sheet =
    withRows(rows)

  def addRow(row : Row) =
    copy(rows = rows ++ List(row))

  def addRows(additionalRows : Iterable[Row]) =
    copy(rows = rows ++ additionalRows)

  def removeRow(row : Row) =
    copy(rows = rows.filter(_ != row))

  def removeRows(whereCondition : Row => Boolean) =
    copy(rows = rows.filter(whereCondition))

  def withMergedRegions(mergedRegions: Iterable[CellRange]): Sheet =
    copy(mergedRegions = mergedRegions.toList)

  def withMergedRegions(mergedRegions: CellRange*): Sheet =
    withMergedRegions(mergedRegions)
  
  def addMergedRegion(mergedRegion : CellRange) =
    copy(mergedRegions = mergedRegions ++ List(mergedRegion))

  def addMergedRegions(additionalMergedRegions : Iterable[CellRange]) =
    copy(mergedRegions = mergedRegions ++ additionalMergedRegions)

  def removeMergedRegion(mergedRegion : CellRange) =
    copy(mergedRegions = mergedRegions.filter(_ != mergedRegion))

  def removeMergedRegions(whereCondition : CellRange => Boolean) =
    copy(mergedRegions = mergedRegions.filter(whereCondition))

  def withPrintSetup(printSetup: PrintSetup) =
    copy(printSetup = Option(printSetup))

  def withoutPrintSetup =
    copy(printSetup = None)

  def withHeader(header: Header) =
    copy(header = Option(header))

  def withoutHeader =
    copy(header = None)

  def withFooter(footer: Footer) =
    copy(footer = Option(footer))

  def withoutFooter =
    copy(footer = None)

  def withMargins(margins: Margins) =
    copy(margins = Option(margins))

  def withoutMargins =
    copy(margins = None)

  def withSplitPane(splitPane: SplitPane) =
    copy(paneAction = Option(splitPane))

  def withFreezePane(freezePane: FreezePane) =
    copy(paneAction = Option(freezePane))

  def withoutSplitOrFreezePane =
    copy(paneAction = Some(NoSplitOrFreeze()))

  def withRepeatingRows(repeatingRows: RowRange) =
    copy(repeatingRows = Option(repeatingRows))

  def withoutRepeatingRows =
    copy(repeatingRows = None)

  def withRepeatingColumns(repeatingColumns: ColumnRange) =
    copy(repeatingColumns = Option(repeatingColumns))

  def withoutRepeatingColumns =
    copy(repeatingColumns = None)

  def withStyle(style : CellStyle) =
    copy(style = Option(style))

  def withoutStyle =
    copy(style = None)
}
