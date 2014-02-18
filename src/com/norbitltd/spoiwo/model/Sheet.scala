package com.norbitltd.spoiwo.model

import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}
import com.norbitltd.spoiwo.natures.csv.CsvProperties

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
            repeatingRows : RowRange = defaultRepeatingRows,
            repeatingColumns : ColumnRange = defaultRepeatingColumns): Sheet =
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

  def apply(name : String, row : Row, rows: Row*) : Sheet = apply(name = name, rows = row :: rows.toList)

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
                          paneAction : Option[PaneAction],
                          repeatingRows : Option[RowRange],
                          repeatingColumns : Option[ColumnRange]) {

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

  def withSplitPane(splitPane : SplitPane) =
    copy(paneAction = Option(splitPane))

  def withFreezePane(freezePane : FreezePane) =
    copy(paneAction = Option(freezePane))

  def withRepeatingRows(repeatingRows : RowRange) =
    copy(repeatingRows = Option(repeatingRows))

  def withRepeatingColumns(repeatingColumns : ColumnRange) =
    copy(repeatingColumns = Option(repeatingColumns))

  def convertToCSV(properties : CsvProperties = CsvProperties.Default) : (String, String) = {
    name.getOrElse("") -> rows.map(r => r.convertToCSV(properties)).mkString("\n")
  }


  def convert(workbook: XSSFWorkbook): XSSFSheet = {
    val sheetName = name.getOrElse("Sheet " + (workbook.getNumberOfSheets + 1))
    val sheet = workbook.createSheet(sheetName)

    updateColumnsWithIndexes().foreach( _.applyTo(sheet))
    rows.foreach(row => row.convertToXLSX(sheet))
    mergedRegions.foreach(mergedRegion => sheet.addMergedRegion(mergedRegion.convert()))

    printSetup.foreach(_.applyTo(sheet))
    header.foreach(_.applyTo(sheet))
    footer.foreach(_.applyTo(sheet))
    properties.foreach(_.applyTo(sheet))
    margins.foreach(_.applyTo(sheet))
    paneAction.foreach(_.applyTo(sheet))

    repeatingRows.foreach(rr => sheet.setRepeatingRows(rr.convert()))
    repeatingColumns.foreach(rc => sheet.setRepeatingColumns(rc.convert()))

    sheet
  }

  def saveAsXlsx(fileName: String) {
    Workbook(this).saveAsXlsx(fileName)
  }

  def saveAsCsv(fileName : String, properties : CsvProperties = CsvProperties.Default) {
    Workbook(this).saveAsCsv(fileName, properties)
  }

  private def updateColumnsWithIndexes(): List[Column] = {
    val currentColumnIndexes = columns.map(_.index).flatten.toSet
    if (currentColumnIndexes.isEmpty) {
      columns.zipWithIndex.map {
        case (column, index) => column.withIndex(index)
      }
    } else if (currentColumnIndexes.size == columns.size) {
      columns
    } else {
      throw new IllegalArgumentException(
        "When explicitly specifying column index you are required to provide it " +
          "uniquely for all columns in this sheet definition!")
    }
  }

}
