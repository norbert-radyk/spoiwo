package com.norbitltd.spoiwo.natures.xlsx

import Model2XlsxConversions.convertWorkbook
import com.norbitltd.spoiwo.model.{Sheet, Workbook}
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.scalatest.{FlatSpec, Matchers}

class Model2XlsxConversionsForWorkbookSpec extends FlatSpec with Matchers {

  def defaultWorkbook = Workbook(
    Sheet(name = "Sheet 1"),
    Sheet(name = "Sheet 2")
  )

  val default: XSSFWorkbook = convertWorkbook(defaultWorkbook)

  "Workbook conversion" should "return a workbook with 1 sheet when default converted" in {
    default.getNumberOfSheets shouldBe 2
    default.getSheetAt(0).getSheetName shouldBe "Sheet 1"
  }

  it should "convert all the sheets provided in order" in {
    val model = Workbook(
      Sheet(name = "Jeden"),
      Sheet(name = "Dwa"),
      Sheet(name = "Trzy")
    )
    val xlsx = convertWorkbook(model)
    xlsx.getNumberOfSheets shouldBe 3
    xlsx.getSheetAt(0).getSheetName shouldBe "Jeden"
    xlsx.getSheetAt(1).getSheetName shouldBe "Dwa"
    xlsx.getSheetAt(2).getSheetName shouldBe "Trzy"
  }

  it should "return first sheet as active sheet by default" in {
    default.getActiveSheetIndex shouldBe 0
  }

  it should "return stated sheet as active sheet when explicitly set" in {
    val model = defaultWorkbook.withActiveSheet(1)
    val xlsx = convertWorkbook(model)
    xlsx.getActiveSheetIndex shouldBe 1
  }

  it should "return first visible tab for first sheet by default" in {
    default.getFirstVisibleTab shouldBe 0
  }

  it should "return first visible tab for the 2nd sheet when set explicitly" in {
    val model = defaultWorkbook.withFirstVisibleTab(1)
    val xlsx = convertWorkbook(model)
    xlsx.getFirstVisibleTab shouldBe 1
  }

}
