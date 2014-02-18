package com.norbitltd.spoiwo.natures.csv

import com.norbitltd.spoiwo.utils.FileUtils
import com.norbitltd.spoiwo.model._

object Model2CsvConversions {

  implicit class CsvWorkbook(wb : Workbook) {

    /**
     * Converts the defined workbook into the sheet name -> csv content map for all of the sheets.
     * @return A sheet name -> CSV content map for each of the sheets
     */
    def convertAsCsv(properties : CsvProperties = CsvProperties.Default) : Map[String, String] =
      convertWorkbookToCsv(wb, properties)

    def saveAsCsv(fileName : String, properties : CsvProperties = CsvProperties.Default) {
      val convertedCsvData = convertAsCsv(properties)
      if( wb.sheets.size <= 1 ) {
        convertedCsvData.values.foreach(csvContent => FileUtils.write(fileName, csvContent))
      } else {
        val fileNameCore = fileName.replace(".csv", "").replace(".CSV", "")
        convertedCsvData.foreach { case( sheetName, csvContent) =>
          val sheetFileName = fileNameCore + "_" + sheetName + ".csv"
          FileUtils.write(sheetFileName, csvContent)
        }
      }
    }
  }

  implicit class CsvSheet(s : Sheet) {

    def convertAsCsv(properties : CsvProperties = CsvProperties.Default) : String =
      convertSheetToCsv(s, properties)._1

    def saveAsCsv(fileName : String, properties : CsvProperties = CsvProperties.Default) {
      Workbook(s).saveAsCsv(fileName, properties)
    }
  }

  private def convertWorkbookToCsv(wb : Workbook, properties : CsvProperties = CsvProperties.Default) : Map[String, String] = {
    require(wb.sheets.size <= 1 || wb.sheets.forall(_.name.isDefined),
      "When converting workbook with multiple sheets to CSV format it is required to specify the unique name for each of them!")
    wb.sheets.map(s => convertSheetToCsv(s, properties)).toMap
  }

  private def convertSheetToCsv(s : Sheet, properties : CsvProperties = CsvProperties.Default) : (String, String) = {
    s.name.getOrElse("") -> s.rows.map(r => convertRowToCsv(r, properties)).mkString("\n")
  }

  private def convertRowToCsv(r : Row, properties : CsvProperties = CsvProperties.Default) : String =
    r.cells.map(c => convertCellToCsv(c, properties)).mkString(properties.separator)

  private def convertCellToCsv(c : Cell, properties : CsvProperties) : String = c match {
    case x : StringCell => x.value
    case x : NumericCell => x.value.toString
    case x : BooleanCell => x.value.toString
    case x : DateCell => x.value.toString
    case x : CalendarCell => x.value.toString
    case x : FormulaCell => throw new IllegalArgumentException("Use of formulas not allowed when converting to CSV format!")
    case x : ErrorValueCell => throw new IllegalArgumentException("Use of error value cells not allowed when converting to CSV format!")
  }
}
