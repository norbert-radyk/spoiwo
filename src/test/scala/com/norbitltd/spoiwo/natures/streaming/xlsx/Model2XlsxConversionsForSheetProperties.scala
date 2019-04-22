package com.norbitltd.spoiwo.natures.streaming.xlsx

import com.norbitltd.spoiwo.model.Height._
import com.norbitltd.spoiwo.model.{CellRange, SheetProperties}
import com.norbitltd.spoiwo.natures.streaming.xlsx.Model2XlsxConversions.convertSheetProperties
import org.apache.poi.ss.util.CellAddress
import org.apache.poi.xssf.streaming.{SXSSFSheet, SXSSFWorkbook}
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.scalatest.{FlatSpec, Matchers}

import scala.language.postfixOps

class Model2XlsxConversionsForSheetProperties extends FlatSpec with Matchers {

  private def apply(properties: SheetProperties): SXSSFSheet = {
    val sheet = new SXSSFWorkbook().createSheet()
    convertSheetProperties(properties, sheet)
    sheet
  }

  private val defaultSheet: SXSSFSheet = apply(SheetProperties())

  "Sheet properties conversion" should "return no active cell by default" in {
    defaultSheet.getActiveCell shouldBe null
  }

  it should "return B3 as active cell if explicitly specified" in {
    val model = SheetProperties(activeCell = "B3")
    val xssf = apply(model)
    xssf.getActiveCell shouldBe new CellAddress("B3")
  }

  it should "return no auto-breaks by default" in {
    defaultSheet.getAutobreaks shouldBe true
  }

  it should "return no auto-break when explicitly set to false" in {
    val model = SheetProperties(autoBreaks = false)
    val xssf = apply(model)
    xssf.getAutobreaks shouldBe false
  }

  it should "return auto-breaks when explicitly set to true" in {
    val model = SheetProperties(autoBreaks = true)
    val xssf = apply(model)
    xssf.getAutobreaks shouldBe true
  }

  it should "return 8 characters width by default" in {
    defaultSheet.getDefaultColumnWidth shouldBe 8
  }

  it should "return 12 characters width when set explicitly" in {
    val model = SheetProperties(defaultColumnWidth = 12)
    val xssf = apply(model)
    xssf.getDefaultColumnWidth shouldBe 12
  }

  it should "return 15 points height by default" in {
    defaultSheet.getDefaultRowHeightInPoints shouldBe 15
    defaultSheet.getDefaultRowHeight shouldBe 300
  }

  it should "return 20 points height when set explicitly" in {
    val model = SheetProperties(defaultRowHeight = 20 points)
    val xssf = apply(model)
    xssf.getDefaultRowHeightInPoints shouldBe 20
    xssf.getDefaultRowHeight shouldBe 400
  }

  it should "return no display formulas by default" in {
    defaultSheet.isDisplayFormulas shouldBe false
  }

  it should "return no display formulas when set to false" in {
    val model = SheetProperties(displayFormulas = false)
    val xssf = apply(model)
    xssf.isDisplayFormulas shouldBe false
  }

  it should "return display formulas when set to true" in {
    val model = SheetProperties(displayFormulas = true)
    val xssf = apply(model)
    xssf.isDisplayFormulas shouldBe true
  }

  it should "return display grid lines by default" in {
    defaultSheet.isDisplayGridlines shouldBe true
  }

  it should "return no display grid lines when set to false" in {
    val model = SheetProperties(displayGridLines = false)
    val xssf = apply(model)
    xssf.isDisplayGridlines shouldBe false
  }

  it should "return display grid lines when set to true" in {
    val model = SheetProperties(displayGridLines = true)
    val xssf = apply(model)
    xssf.isDisplayGridlines shouldBe true
  }

  it should "return display guts by default" in {
    defaultSheet.getDisplayGuts shouldBe true
  }

  it should "return no display guts when set to false" in {
    val model = SheetProperties(displayGuts = false)
    val xssf = apply(model)
    xssf.getDisplayGuts shouldBe false
  }

  it should "return display guts when set to true" in {
    val model = SheetProperties(displayGuts = true)
    val xssf = apply(model)
    xssf.getDisplayGuts shouldBe true
  }

  it should "return display row/column headings by default" in {
    defaultSheet.isDisplayRowColHeadings shouldBe true
  }

  it should "return no display row/column headings when set to false" in {
    val model = SheetProperties(displayRowColHeadings = false)
    val xssf = apply(model)
    xssf.isDisplayRowColHeadings shouldBe false
  }

  it should "return display row/column headings when set to true" in {
    val model = SheetProperties(displayRowColHeadings = true)
    val xssf = apply(model)
    xssf.isDisplayRowColHeadings shouldBe true
  }

  it should "return display zeros by default" in {
    defaultSheet.isDisplayZeros shouldBe true
  }

  it should "return no display zeros when set to false" in {
    val model = SheetProperties(displayZeros = false)
    val xssf = apply(model)
    xssf.isDisplayZeros shouldBe false
  }

  it should "return display zeros when set to true" in {
    val model = SheetProperties(displayZeros = true)
    val xssf = apply(model)
    xssf.isDisplayZeros shouldBe true
  }

  it should "not fit sheet to page by default" in {
    defaultSheet.getFitToPage shouldBe false
  }

  it should "not fit sheet to page when set to false" in {
    val model = SheetProperties(fitToPage = false)
    val xssf = apply(model)
    xssf.getFitToPage shouldBe false
  }

  it should "fit sheet to page when set to true" in {
    val model = SheetProperties(fitToPage = true)
    val xssf = apply(model)
    xssf.getFitToPage shouldBe true
  }

  it should "not force formula recalculation by default" in {
    defaultSheet.getForceFormulaRecalculation shouldBe false
  }

  it should "not force formula recalculation when set to false" in {
    val model = SheetProperties(forceFormulaRecalculation = false)
    val xssf = apply(model)
    xssf.getForceFormulaRecalculation shouldBe false
  }

  it should "force formula recalculation when set to true" in {
    val model = SheetProperties(forceFormulaRecalculation = true)
    val xssf = apply(model)
    xssf.getForceFormulaRecalculation shouldBe true
  }

  it should "not get horizontally center by default" in {
    defaultSheet.getHorizontallyCenter shouldBe false
  }

  it should "not get horizontally center when set to false" in {
    val model = SheetProperties(horizontallyCenter = false)
    val xssf = apply(model)
    xssf.getHorizontallyCenter shouldBe false
  }

  it should "get horizontally center when set to true" in {
    val model = SheetProperties(horizontallyCenter = true)
    val xssf = apply(model)
    xssf.getHorizontallyCenter shouldBe true
  }

  it should "not return the print area by default" in {
    defaultSheet.getWorkbook.getPrintArea(0) shouldBe null
  }

  it should "return the print area  when set explicitly" in {
    val model = SheetProperties(printArea = CellRange(2 -> 7, 1 -> 4))
    val xssf = apply(model)
    xssf.getWorkbook.getPrintArea(0) shouldBe "Sheet0!$B$3:$E$8"
  }

  it should "print grid lines by default" in {
    defaultSheet.isPrintGridlines shouldBe false
  }

  it should "not print grid lines when set to false" in {
    val model = SheetProperties(printGridLines = false)
    val xssf = apply(model)
    xssf.isPrintGridlines shouldBe false
  }

  it should "get print grid lines when set to true" in {
    val model = SheetProperties(printGridLines = true)
    val xssf = apply(model)
    xssf.isPrintGridlines shouldBe true
  }

  it should "not be right to left by default" in {
    defaultSheet.isRightToLeft shouldBe false
  }

  it should "not be right to left when set to false" in {
    val model = SheetProperties(rightToLeft = false)
    val xssf = apply(model)
    xssf.isRightToLeft shouldBe false
  }

  it should "be right to left when set to true" in {
    val model = SheetProperties(rightToLeft = true)
    val xssf = apply(model)
    xssf.isRightToLeft shouldBe true
  }

  it should "do row sums below by default" in {
    defaultSheet.getRowSumsBelow shouldBe true
  }

  it should "not do row sums below when set to false" in {
    val model = SheetProperties(rowSumsBelow = false)
    val xssf = apply(model)
    xssf.getRowSumsBelow shouldBe false
  }

  it should "be do row sums below when set to true" in {
    val model = SheetProperties(rowSumsBelow = true)
    val xssf = apply(model)
    xssf.getRowSumsBelow shouldBe true
  }

  it should "do row sums right by default" in {
    defaultSheet.getRowSumsRight shouldBe true
  }

  it should "not do row sums right when set to false" in {
    val model = SheetProperties(rowSumsRight = false)
    val xssf = apply(model)
    xssf.getRowSumsRight shouldBe false
  }

  it should "be do row sums right when set to true" in {
    val model = SheetProperties(rowSumsRight = true)
    val xssf = apply(model)
    xssf.getRowSumsRight shouldBe true
  }

  it should "be selected by default when only 1 sheet" in {
    defaultSheet.isSelected shouldBe true
  }

  it should "not be selected when set to false, but only one sheet" in {
    val model = SheetProperties(selected = false)
    val xssf = apply(model)
    xssf.isSelected shouldBe false
  }

  it should "not be selected when set to false, but there are multiple sheets" in {
    val workbook = new XSSFWorkbook()
    val sheet1 = workbook.createSheet("S1")
    val sheet2 = workbook.createSheet("S2")
    val model1 = SheetProperties(selected = false)
    val model2 = SheetProperties(selected = true)
    convertSheetProperties(model1, sheet1)
    convertSheetProperties(model2, sheet2)
    sheet1.isSelected shouldBe false
    sheet2.isSelected shouldBe true
  }

  it should "be selected when set to true" in {
    val model = SheetProperties(selected = true)
    val xssf = apply(model)
    xssf.isSelected shouldBe true
  }

  it should "not be virtually center by default" in {
    defaultSheet.getVerticallyCenter shouldBe false
  }

  it should "not be virtually center when set to false" in {
    val model = SheetProperties(virtuallyCenter = false)
    val xssf = apply(model)
    xssf.getVerticallyCenter shouldBe false
  }

  it should "be virtually center when set to true" in {
    val model = SheetProperties(virtuallyCenter = true)
    val xssf = apply(model)
    xssf.getVerticallyCenter shouldBe true
  }

}
