package spoiwo.natures.csv

import com.github.tototoshi.csv.{CSVWriter, DefaultCSVFormat, defaultCSVFormat}
import spoiwo.utils.FileUtils
import spoiwo.model._

import java.io.StringWriter
import java.time.ZoneId

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

  private val customCsvFormat = new DefaultCSVFormat {
    override val lineTerminator: String = "\n"
  }

  private def convertSheetToCsv(s: Sheet, properties: CsvProperties): (String, String) = {
    val format = if (properties.separator.isEmpty || properties.separator == ",") {
      customCsvFormat
    } else {
      new DefaultCSVFormat {
        override val delimiter: Char = properties.separator.charAt(0)
        override val lineTerminator: String = customCsvFormat.lineTerminator
      }
    }
    val sw = new StringWriter()
    try {
      val csvWriter = CSVWriter.open(sw)(format)
      try {
        s.rows.map(r => writeRowAsCsv(csvWriter, r, properties))
      } finally {
        csvWriter.close()
      }
    } finally {
      sw.close()
    }
    s.name.getOrElse("") -> sw.toString
  }

  private def writeRowAsCsv(csvWriter: CSVWriter, r: Row, properties: CsvProperties): Unit = {
    csvWriter.writeRow(r.cells.map(c => convertCellToCsv(c, properties)).toSeq)
  }

  private def convertCellToCsv(c: Cell, properties: CsvProperties): String = c match {
    case _: BlankCell     => ""
    case x: StringCell    => x.value
    case x: NumericCell   => x.value.toString
    case x: BooleanCell   => if (x.value) properties.defaultBooleanTrueString else properties.defaultBooleanFalseString
    case x: DateCell      => x.value.toInstant.atZone(ZoneId.systemDefault()).format(properties.defaultDateFormatter)
    case x: CalendarCell  => x.value.toInstant.atZone(ZoneId.systemDefault()).format(properties.defaultDateFormatter)
    case x: HyperLinkCell => x.value.text
    case value            => throw new IllegalArgumentException(s"Unable to convert to CSV cell for value: $value!")
  }
}
