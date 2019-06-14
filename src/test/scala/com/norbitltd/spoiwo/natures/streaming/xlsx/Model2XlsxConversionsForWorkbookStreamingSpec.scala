package com.norbitltd.spoiwo.natures.streaming.xlsx

import com.norbitltd.spoiwo.model.{Sheet, Workbook}
import com.norbitltd.spoiwo.natures.streaming.xlsx.Model2XlsxConversions.{convertWorkbook, writeToExistingWorkbook}
import com.norbitltd.spoiwo.natures.xlsx.Utils._
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.scalatest.{FlatSpec, Matchers}

class Model2XlsxConversionsForWorkbookStreamingSpec extends FlatSpec with Matchers {

  def defaultWorkbook = Workbook(
    Sheet(name = "Sheet 1"),
    Sheet(name = "Sheet 2")
  )

  val default: SXSSFWorkbook = convertWorkbook(defaultWorkbook)

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

  it should "be able to merge sheets into existing workbook" in {
    val previousSheet = generateSheet(0 to 2, 0 to 5).withSheetName("Overwrite")
    val untouchedSheet = generateSheet(0 to 3, 0 to 5).withSheetName("Untouched")
    val previousModel = Workbook(previousSheet, untouchedSheet)
    val existingWorkbook = convertWorkbook(previousModel)
    val overwritingSheet =
      generateSheet[Int, Int](1 to 3, 1 to 4, (rowNum, colNum) => s"NEW $colNum,$rowNum").withSheetName("Overwrite")
    val newSheet = generateSheet(0 to 3, 0 to 5).withSheetName("New")
    val newModel = Workbook(
      overwritingSheet,
      newSheet
    )
    writeToExistingWorkbook(newModel, existingWorkbook)
    existingWorkbook.getNumberOfSheets shouldBe 3
    existingWorkbook.getSheetAt(0).getSheetName shouldBe "Overwrite"
    existingWorkbook.getSheetAt(1).getSheetName shouldBe "Untouched"
    existingWorkbook.getSheetAt(2).getSheetName shouldBe "New"
    val overwrittenPoiSheet = existingWorkbook.getSheet("Overwrite")
    val overwrittenSheetData = mergeSheetData(Seq(previousSheet, overwritingSheet), _.value.toString)
    nonEqualCells(overwrittenPoiSheet, overwrittenSheetData, _.getStringCellValue) shouldBe empty
    val untouchedSheetData = mergeSheetData(Seq(untouchedSheet), _.value.toString)
    nonEqualCells(existingWorkbook.getSheet("Untouched"), untouchedSheetData, _.getStringCellValue) shouldBe empty
    val newSheetData = mergeSheetData(Seq(newSheet), _.value.toString)
    nonEqualCells(existingWorkbook.getSheet("New"), newSheetData, _.getStringCellValue) shouldBe empty
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
