package com.norbitltd.spoiwo.natures.csv

import com.norbitltd.spoiwo.utils.FileUtils
import com.norbitltd.spoiwo.model._
import org.joda.time.LocalDate

object Model2CsvConversions {

  implicit class CsvWorkbook(wb: Workbook) {

    /**
      * Converts the defined workbook into the sheet name -> csv content map for all of the sheets.
      * @return A sheet name -> CSV content map for each of the sheets
      */
    def convertAsCsv(properties: CsvProperties = CsvProperties.Default): Map[String, String] =
      convertWorkbookToCsv(wb, properties)

    def saveAsCsv(fileName: String, properties: CsvProperties = CsvProperties.Default): Unit = {
      val convertedCsvData = convertAsCsv(properties)
      if (wb.sheets.size <= 1) {
        convertedCsvData.values.foreach(csvContent => FileUtils.write(fileName, csvContent))
      } else {
        val fileNameCore = fileName.replace(".csv", "").replace(".CSV", "")
        convertedCsvData.foreach {
          case (sheetName, csvContent) =>
            val sheetFileName = fileNameCore + "_" + sheetName + ".csv"
            FileUtils.write(sheetFileName, csvContent)
        }
      }
    }
  }

  implicit class CsvSheet(s: Sheet) {

    def convertAsCsv(properties: CsvProperties = CsvProperties.Default): String =
      convertSheetToCsv(s, properties)._2

    def saveAsCsv(fileName: String, properties: CsvProperties = CsvProperties.Default): Unit = {
      Workbook(s).saveAsCsv(fileName, properties)
    }
  }

  private def convertWorkbookToCsv(wb: Workbook, properties: CsvProperties): Map[String, String] = {
    require(
      wb.sheets.size <= 1 || wb.sheets.forall(_.name.isDefined),
      "When converting workbook with multiple sheets to CSV format it is required to specify the unique name for each of them!"
    )
    wb.sheets.map(s => convertSheetToCsv(s, properties)).toMap
  }

  private def convertSheetToCsv(s: Sheet, properties: CsvProperties): (String, String) = {
    s.name.getOrElse("") -> s.rows.map(r => convertRowToCsv(r, properties)).mkString("\n")
  }

  private def convertRowToCsv(r: Row, properties: CsvProperties): String =
    r.cells.map(c => convertCellToCsv(c, properties)).mkString(properties.separator)

  private def convertCellToCsv(c: Cell, properties: CsvProperties): String = c match {
    case _: BlankCell    => ""
    case x: StringCell   => x.value
    case x: NumericCell  => x.value.toString
    case x: BooleanCell  => if (x.value) properties.defaultBooleanTrueString else properties.defaultBooleanFalseString
    case x: DateCell     => LocalDate.fromDateFields(x.value).toString(properties.defaultDateFormat)
    case x: CalendarCell => LocalDate.fromCalendarFields(x.value).toString(properties.defaultDateFormat)
    case x: HyperLinkUrlCell => x.value.text
    case value           => throw new IllegalArgumentException(s"Unable to convert to CSV cell for value: $value!")
  }
}
