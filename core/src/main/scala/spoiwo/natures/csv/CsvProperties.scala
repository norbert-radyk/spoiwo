package spoiwo.natures.csv

import java.time.format.DateTimeFormatter

object CsvProperties {

  val Default: CsvProperties = CsvProperties()

}

case class CsvProperties(
    separator: Char = ',',
    defaultDateFormat: String = "yyyy-MM-dd",
    defaultBooleanTrueString: String = "true",
    defaultBooleanFalseString: String = "false"
) {

  val defaultDateFormatter: DateTimeFormatter = DateTimeFormatter.ofPattern(defaultDateFormat)

}
