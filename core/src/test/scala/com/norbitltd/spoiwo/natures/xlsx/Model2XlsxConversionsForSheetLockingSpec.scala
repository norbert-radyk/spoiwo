package com.norbitltd.spoiwo.natures.xlsx

import com.norbitltd.spoiwo.model.SheetLocking
import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class Model2XlsxConversionsForSheetLockingSpec extends AnyFlatSpec with Matchers {

  private def apply(locking: SheetLocking): XSSFSheet = {
    val sheet = new XSSFWorkbook().createSheet()
    Model2XlsxConversions.convertSheetLocking(locking, sheet)
    sheet
  }

  "Sheet locking" should "not be enabled if locking has not been applied" in {
    val sheet = new XSSFWorkbook().createSheet()
    sheet.isSheetLocked shouldBe false
  }

  it should "enable sheet locking when sheet locking properties specified" in {
    val sheetLocking = SheetLocking(lockedAutoFilter = true)
    val sheet = apply(sheetLocking)
    sheet.isSheetLocked shouldBe true
  }

}
