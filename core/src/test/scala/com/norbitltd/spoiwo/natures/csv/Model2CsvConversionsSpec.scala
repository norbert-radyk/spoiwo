package com.norbitltd.spoiwo.natures.csv

import java.time.LocalDate

import com.norbitltd.spoiwo.model.{HyperLink, Row, Sheet}
import Model2CsvConversions._
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class Model2CsvConversionsSpec extends AnyFlatSpec with Matchers {

  "Model to CSV conversion" should "correctly convert the single text-only sheet" in {
    val sheet = Sheet(name = "CSV conversion").withRows(
      Row().withCellValues("EUROPE", "Poland", "Wroclaw"),
      Row().withCellValues("EUROPE", "United Kingdom", "London"),
      Row().withCellValues("ASIA", "China", "Tianjin")
    )

    val csvText = "EUROPE,Poland,Wroclaw\nEUROPE,United Kingdom,London\nASIA,China,Tianjin"
    csvText shouldBe sheet.convertAsCsv()
  }

  it should "correctly convert the single text-only sheet with '|' separator" in {
    val sheet = Sheet(name = "CSV conversion").withRows(
      Row().withCellValues("EUROPE", "Poland", "Wroclaw"),
      Row().withCellValues("EUROPE", "United Kingdom", "London"),
      Row().withCellValues("ASIA", "China", "Tianjin")
    )
    val properties = CsvProperties(separator = "|")

    val csvText = "EUROPE|Poland|Wroclaw\nEUROPE|United Kingdom|London\nASIA|China|Tianjin"
    csvText shouldBe sheet.convertAsCsv(properties)
  }

  it should "correctly format boolean value with default 'true' and 'false'" in {
    val sheet = Sheet(name = "CSV conversion").withRows(
      Row().withCellValues("City", "Is Capital"),
      Row().withCellValues("London", true),
      Row().withCellValues("Wroclaw", false)
    )

    val csvText = "City,Is Capital\nLondon,true\nWroclaw,false"
    csvText shouldBe sheet.convertAsCsv()
  }

  it should "correctly format boolean value with 'Y' and 'N'" in {
    val sheet = Sheet(name = "CSV conversion").withRows(
      Row().withCellValues("City", "Is Capital"),
      Row().withCellValues("London", true),
      Row().withCellValues("Wroclaw", false)
    )
    val properties = CsvProperties(defaultBooleanTrueString = "Y", defaultBooleanFalseString = "N")

    val csvText = "City,Is Capital\nLondon,Y\nWroclaw,N"
    csvText shouldBe sheet.convertAsCsv(properties)
  }

  it should "correctly format date value to yyyy-MM-dd by default" in {
    val sheet = Sheet(name = "CSV conversion").withRows(
      Row().withCellValues("Albert Einstein", LocalDate.of(1879, 3, 14)),
      Row().withCellValues("Thomas Edison", LocalDate.of(1847, 2, 11))
    )

    val csvText = "Albert Einstein,1879-03-14\nThomas Edison,1847-02-11"
    csvText shouldBe sheet.convertAsCsv()
  }

  it should "correctly format date value to yyyy/MM/dd by default" in {
    val sheet = Sheet(name = "CSV conversion").withRows(
      Row().withCellValues("Albert Einstein", LocalDate.of(1879, 3, 14)),
      Row().withCellValues("Thomas Edison", LocalDate.of(1847, 2, 11))
    )
    val properties = CsvProperties(defaultDateFormat = "yyyy/MM/dd")

    val csvText = "Albert Einstein,1879/03/14\nThomas Edison,1847/02/11"
    csvText shouldBe sheet.convertAsCsv(properties)
  }

  it should "correctly process the HyperLinkUrl values" in {
    val sheet = Sheet(name = "CSV conversion").withRows(
      Row().withCellValues(
        HyperLink(text = "spoiwo issue", address = "https://github.com/norbert-radyk/spoiwo/issues/34")
      )
    )

    val csvText = "spoiwo issue"
    csvText shouldBe sheet.convertAsCsv()
  }

}
