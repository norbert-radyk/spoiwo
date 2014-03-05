package com.norbitltd.spoiwo.model

object Sheet extends Factory {

  private lazy val defaultName = ""
  private lazy val defaultColumns = Nil
  private lazy val defaultRows = Nil
  private lazy val defaultMergedRegions = Nil
  private lazy val defaultPrintSetup = PrintSetup.Default
  private lazy val defaultHeader = Header.Empty
  private lazy val defaultFooter = Footer.Empty
  private lazy val defaultProperties = SheetProperties.Default
  private lazy val defaultMargins = Margins.Default
  private lazy val defaultPaneAction = NoSplitOrFreeze()
  private lazy val defaultRepeatingRows = RowRange.None
  private lazy val defaultRepeatingColumns = ColumnRange.None

  val Blank = Sheet()

  def apply(name: String = defaultName,
            columns: List[Column] = defaultColumns,
            rows: List[Row] = defaultRows,
            mergedRegions: List[CellRange] = defaultMergedRegions,
            printSetup: PrintSetup = defaultPrintSetup,
            header: Header = defaultHeader,
            footer: Footer = defaultFooter,
            properties: SheetProperties = defaultProperties,
            margins: Margins = defaultMargins,
            paneAction: PaneAction = defaultPaneAction,
            repeatingRows: RowRange = defaultRepeatingRows,
            repeatingColumns: ColumnRange = defaultRepeatingColumns): Sheet =
    apply(
      name = wrap(name, defaultName),
      columns = columns,
      rows = rows,
      mergedRegions = mergedRegions,
      printSetup = wrap(printSetup, defaultPrintSetup),
      header = wrap(header, defaultHeader),
      footer = wrap(footer, defaultFooter),
      properties = wrap(properties, defaultProperties),
      margins = wrap(margins, defaultMargins),
      paneAction = wrap(paneAction, defaultPaneAction),
      repeatingRows = wrap(repeatingRows, defaultRepeatingRows),
      repeatingColumns = wrap(repeatingColumns, defaultRepeatingColumns)
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
                          repeatingColumns: Option[ColumnRange]) {

  def withSheetName(name: String) =
    copy(name = Option(name))

  def withColumns(columns: List[Column]): Sheet =
    copy(columns = columns)

  def withColumns(columns: Column*): Sheet =
    withColumns(columns.toList)

  def withRows(rows: Iterable[Row]): Sheet =
    copy(rows = rows.toList)

  def withRows(rows: Row*): Sheet =
    withRows(rows)

  def withMergedRegions(mergedRegions: Iterable[CellRange]): Sheet =
    copy(mergedRegions = mergedRegions.toList)

  def withMergedRegions(mergedRegions: CellRange*): Sheet =
    withMergedRegions(mergedRegions)

  def withPrintSetup(printSetup: PrintSetup) =
    copy(printSetup = Option(printSetup))

  def withHeader(header: Header) =
    copy(header = Option(header))

  def withFooter(footer: Footer) =
    copy(footer = Option(footer))

  def withMargins(margins: Margins) =
    copy(margins = Option(margins))

  def withSplitPane(splitPane: SplitPane) =
    copy(paneAction = Option(splitPane))

  def withFreezePane(freezePane: FreezePane) =
    copy(paneAction = Option(freezePane))

  def withRepeatingRows(repeatingRows: RowRange) =
    copy(repeatingRows = Option(repeatingRows))

  def withRepeatingColumns(repeatingColumns: ColumnRange) =
    copy(repeatingColumns = Option(repeatingColumns))
}
