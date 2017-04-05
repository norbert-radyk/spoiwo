package com.norbitltd.spoiwo.natures.xlsx

import Model2XlsxConversions.convertWorkbook
import com.norbitltd.spoiwo.model.{Sheet, Workbook}
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.scalatest.FlatSpec
import org.scalatest.Matchers._

class Model2XlsxConversionsForWorkbookSpec extends FlatSpec {

  def defaultWorkbook = Workbook(
    Sheet(name = "Sheet 1"),
    Sheet(name = "Sheet 2")
  )

  val default: XSSFWorkbook = convertWorkbook(defaultWorkbook)

  "Workbook conversion" should "return a workbook with 1 sheet when default converted" in {
    assert(default.getNumberOfSheets == 2)
    assert(default.getSheetAt(0).getSheetName == "Sheet 1")
  }

  it should "convert all the sheets provided in order" in {
    val model = Workbook(
      Sheet(name = "Jeden"),
      Sheet(name = "Dwa"),
      Sheet(name = "Trzy")
    )
    val xlsx = convertWorkbook(model)
    assert(xlsx.getNumberOfSheets == 3)
    assert(xlsx.getSheetAt(0).getSheetName == "Jeden")
    assert(xlsx.getSheetAt(1).getSheetName == "Dwa")
    assert(xlsx.getSheetAt(2).getSheetName == "Trzy")
  }

  it should "return first sheet as active sheet by default" in {
    assert(default.getActiveSheetIndex == 0)
  }

  it should "return stated sheet as active sheet when explicitly set" in {
    val model = defaultWorkbook.withActiveSheet(1)
    val xlsx = convertWorkbook(model)
    assert(xlsx.getActiveSheetIndex == 1)
  }

  it should "return first visible tab for first sheet by default" in {
    assert(default.getFirstVisibleTab == 0)
  }

  it should "return first visible tab for the 2nd sheet when set explicitly" in {
    val model = defaultWorkbook.withFirstVisibleTab(1)
    val xlsx = convertWorkbook(model)
    assert(xlsx.getFirstVisibleTab == 1)
  }


}
