package com.norbitltd.spoiwo.natures.xlsx

import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}
import org.apache.poi.ss.util.CellAddress
import org.scalatest.FlatSpec
import com.norbitltd.spoiwo.model.{CellRange, SheetProperties}
import Model2XlsxConversions.convertSheetProperties
import com.norbitltd.spoiwo.model.Height._
import scala.language.postfixOps

class Model2XlsxConversionsForSheetProperties extends FlatSpec {

  private def apply(properties : SheetProperties) : XSSFSheet = {
    val sheet = new XSSFWorkbook().createSheet()
    convertSheetProperties(properties, sheet)
    sheet
  }

  private val defaultSheet : XSSFSheet = apply(SheetProperties())

  "Sheet properties conversion" should "return no auto-filter by default" in {
    //FIXME Unable to test this one as auto-filter value can't be obtained
  }

  it should "return no active cell by default" in {
    assert(defaultSheet.getActiveCell == null)
  }

  it should "return B3 as active cell if explicitly specified" in {
    val model = SheetProperties(activeCell = "B3")
    val xssf = apply(model)
    assert(xssf.getActiveCell == new CellAddress("B3"))
  }

  it should "return no auto-breaks by default" in {
    assert(defaultSheet.getAutobreaks)
  }

  it should "return no auto-break when explicitly set to false" in {
    val model = SheetProperties(autoBreaks = false)
    val xssf = apply(model)
    assert(!xssf.getAutobreaks)
  }

  it should "return auto-breaks when explicitly set to true" in {
    val model = SheetProperties(autoBreaks = true)
    val xssf = apply(model)
    assert(xssf.getAutobreaks)
  }

  it should "return 8 characters width by default" in {
    assert(defaultSheet.getDefaultColumnWidth == 8)
  }

  it should "return 12 characters width when set explicitly" in {
    val model = SheetProperties(defaultColumnWidth = 12)
    val xssf = apply(model)
    assert(xssf.getDefaultColumnWidth == 12)
  }

  it should "return 15 points height by default" in {
    assert(defaultSheet.getDefaultRowHeightInPoints == 15)
    assert(defaultSheet.getDefaultRowHeight == 300)
  }

  it should "return 20 points height when set explicitly" in {
    val model = SheetProperties(defaultRowHeight = 20 points)
    val xssf = apply(model)
    assert(xssf.getDefaultRowHeightInPoints == 20)
    assert(xssf.getDefaultRowHeight == 400)
  }

  it should "return no display formulas by default" in {
    assert(!defaultSheet.isDisplayFormulas)
  }

  it should "return no display formulas when set to false" in {
    val model = SheetProperties(displayFormulas = false)
    val xssf = apply(model)
    assert(!xssf.isDisplayFormulas)
  }

  it should "return display formulas when set to true" in {
    val model = SheetProperties(displayFormulas = true)
    val xssf = apply(model)
    assert(xssf.isDisplayFormulas)
  }

  it should "return display grid lines by default" in {
    assert(defaultSheet.isDisplayGridlines)
  }

  it should "return no display grid lines when set to false" in {
    val model = SheetProperties(displayGridLines = false)
    val xssf = apply(model)
    assert(!xssf.isDisplayGridlines)
  }

  it should "return display grid lines when set to true" in {
    val model = SheetProperties(displayGridLines = true)
    val xssf = apply(model)
    assert(xssf.isDisplayGridlines)
  }

  it should "return display guts by default" in {
    assert(defaultSheet.getDisplayGuts)
  }

  it should "return no display guts when set to false" in {
    val model = SheetProperties(displayGuts = false)
    val xssf = apply(model)
    assert(!xssf.getDisplayGuts)
  }

  it should "return display guts when set to true" in {
    val model = SheetProperties(displayGuts = true)
    val xssf = apply(model)
    assert(xssf.getDisplayGuts)
  }

  it should "return display row/column headings by default" in {
    assert(defaultSheet.isDisplayRowColHeadings)
  }

  it should "return no display row/column headings when set to false" in {
    val model = SheetProperties(displayRowColHeadings = false)
    val xssf = apply(model)
    assert(!xssf.isDisplayRowColHeadings)
  }

  it should "return display row/column headings when set to true" in {
    val model = SheetProperties(displayRowColHeadings = true)
    val xssf = apply(model)
    assert(xssf.isDisplayRowColHeadings)
  }

  it should "return display zeros by default" in {
    assert(defaultSheet.isDisplayZeros)
  }

  it should "return no display zeros when set to false" in {
    val model = SheetProperties(displayZeros = false)
    val xssf = apply(model)
    assert(!xssf.isDisplayZeros)
  }

  it should "return display zeros when set to true" in {
    val model = SheetProperties(displayZeros = true)
    val xssf = apply(model)
    assert(xssf.isDisplayZeros)
  }

  it should "not fit sheet to page by default" in {
    assert(!defaultSheet.getFitToPage)
  }

  it should "not fit sheet to page when set to false" in {
    val model = SheetProperties(fitToPage = false)
    val xssf = apply(model)
    assert(!xssf.getFitToPage)
  }

  it should "fit sheet to page when set to true" in {
    val model = SheetProperties(fitToPage = true)
    val xssf = apply(model)
    assert(xssf.getFitToPage)
  }

  it should "not force formula recalculation by default" in {
    assert(!defaultSheet.getForceFormulaRecalculation)
  }

  it should "not force formula recalculation when set to false" in {
    val model = SheetProperties(forceFormulaRecalculation = false)
    val xssf = apply(model)
    assert(!xssf.getForceFormulaRecalculation)
  }

  it should "force formula recalculation when set to true" in {
    val model = SheetProperties(forceFormulaRecalculation = true)
    val xssf = apply(model)
    assert(xssf.getForceFormulaRecalculation)
  }

  it should "not get horizontally center by default" in {
    assert(!defaultSheet.getHorizontallyCenter)
  }

  it should "not get horizontally center when set to false" in {
    val model = SheetProperties(horizontallyCenter = false)
    val xssf = apply(model)
    assert(!xssf.getHorizontallyCenter)
  }

  it should "get horizontally center when set to true" in {
    val model = SheetProperties(horizontallyCenter = true)
    val xssf = apply(model)
    assert(xssf.getHorizontallyCenter)
  }

  it should "not return the print area by default" in {
    assert(defaultSheet.getWorkbook.getPrintArea(0) == null)
  }

  it should "return the print area  when set explicitly" in {
    val model = SheetProperties(printArea = CellRange(2->7, 1 -> 4))
    val xssf = apply(model)
    assert(xssf.getWorkbook.getPrintArea(0) == "Sheet0!$B$3:$E$8")
  }

  it should "print grid lines by default" in {
    assert(!defaultSheet.isPrintGridlines)
  }

  it should "not print grid lines when set to false" in {
    val model = SheetProperties(printGridLines = false)
    val xssf = apply(model)
    assert(!xssf.isPrintGridlines)
  }

  it should "get print grid lines when set to true" in {
    val model = SheetProperties(printGridLines = true)
    val xssf = apply(model)
    assert(xssf.isPrintGridlines)
  }

  it should "not be right to left by default" in {
    assert(!defaultSheet.isRightToLeft)
  }

  it should "not be right to left when set to false" in {
    val model = SheetProperties(rightToLeft = false)
    val xssf = apply(model)
    assert(!xssf.isRightToLeft)
  }

  it should "be right to left when set to true" in {
    val model = SheetProperties(rightToLeft = true)
    val xssf = apply(model)
    assert(xssf.isRightToLeft)
  }

  it should "do row sums below by default" in {
    assert(defaultSheet.getRowSumsBelow)
  }

  it should "not do row sums below when set to false" in {
    val model = SheetProperties(rowSumsBelow = false)
    val xssf = apply(model)
    assert(!xssf.getRowSumsBelow)
  }

  it should "be do row sums below when set to true" in {
    val model = SheetProperties(rowSumsBelow = true)
    val xssf = apply(model)
    assert(xssf.getRowSumsBelow)
  }

  it should "do row sums right by default" in {
    assert(defaultSheet.getRowSumsRight)
  }

  it should "not do row sums right when set to false" in {
    val model = SheetProperties(rowSumsRight = false)
    val xssf = apply(model)
    assert(!xssf.getRowSumsRight)
  }

  it should "be do row sums right when set to true" in {
    val model = SheetProperties(rowSumsRight = true)
    val xssf = apply(model)
    assert(xssf.getRowSumsRight)
  }

  it should "be selected by default when only 1 sheet" in {
    assert(defaultSheet.isSelected)
  }

  it should "not be selected when set to false, but only one sheet" in {
    val model = SheetProperties(selected = false)
    val xssf = apply(model)
    assert(!xssf.isSelected)
  }

  it should "not be selected when set to false, but there are multiple sheets" in {
    val workbook = new XSSFWorkbook()
    val sheet1 = workbook.createSheet("S1")
    val sheet2 = workbook.createSheet("S2")
    val model1 = SheetProperties(selected = false)
    val model2 = SheetProperties(selected = true)
    convertSheetProperties(model1, sheet1)
    convertSheetProperties(model2, sheet2)
    assert(!sheet1.isSelected)
    assert(sheet2.isSelected)
  }

  it should "be selected when set to true" in {
    val model = SheetProperties(selected = true)
    val xssf = apply(model)
    assert(xssf.isSelected)
  }

  it should "return blank tab color by default" in {
    //FIXME Unable to test - no TabColor method
  }

  it should "not be virtually center by default" in {
    assert(!defaultSheet.getVerticallyCenter)
  }

  it should "not be virtually center when set to false" in {
    val model = SheetProperties(virtuallyCenter = false)
    val xssf = apply(model)
    assert(!xssf.getVerticallyCenter)
  }

  it should "be virtually center when set to true" in {
    val model = SheetProperties(virtuallyCenter = true)
    val xssf = apply(model)
    assert(xssf.getVerticallyCenter)
  }

  it should "have 100% zoom by default" in {
    //FIXME Unable to test - no getZoom method
  }


}
