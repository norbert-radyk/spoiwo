package com.norbitltd.spoiwo.natures.csv

import java.time.format.DateTimeFormatter

object CsvProperties {

  val Default = CsvProperties()

}

case class CsvProperties(separator: String = ",",
                         defaultDateFormat: String = "yyyy-MM-dd",
                         defaultBooleanTrueString: String = "true",
                         defaultBooleanFalseString: String = "false") {

  val defaultDateFormatter: DateTimeFormatter = DateTimeFormatter.ofPattern(defaultDateFormat)

}
