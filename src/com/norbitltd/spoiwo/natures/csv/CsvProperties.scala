package com.norbitltd.spoiwo.natures.csv

object CsvProperties {

  val Default = CsvProperties()

}

case class CsvProperties(separator: String = ",",
                         defaultDateFormat: String = "yyyy-MM-dd",
                         defaultBooleanTrueString: String = "true",
                         defaultBooleanFalseString: String = "false")
