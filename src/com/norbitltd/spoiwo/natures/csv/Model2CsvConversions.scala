package com.norbitltd.spoiwo.natures.csv

import com.norbitltd.spoiwo.utils.FileUtils
import com.norbitltd.spoiwo.model.{Workbook, Sheet}

object Model2CsvConversions {

  implicit class CsvWorkbook(wb : Workbook) {

    /**
     * Converts the defined workbook into the sheet name -> csv content map for all of the sheets.
     * @return A sheet name -> CSV content map for each of the sheets
     */
    def asCsvStrings(properties : CsvProperties = CsvProperties.Default) : Map[String, String] =
      convertWorkbookToCsv(wb, properties)

    def saveAsCsv(fileName : String, properties : CsvProperties = CsvProperties.Default) {
      val convertedCsvData = asCsvStrings(properties)
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

    def asCsvString(properties : CsvProperties = CsvProperties.Default) : String =
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
    s.name.getOrElse("") -> s.rows.map(r => r.convertToCSV(properties)).mkString("\n")
  }

}
