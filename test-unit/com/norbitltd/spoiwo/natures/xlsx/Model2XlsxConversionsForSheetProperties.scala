package com.norbitltd.spoiwo.natures.xlsx

import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFWorkbook}
import org.scalatest.FlatSpec
import com.norbitltd.spoiwo.model.SheetProperties
import Model2XlsxConversions.convertSheetProperties
import com.norbitltd.spoiwo.model.Measure._

class Model2XlsxConversionsForSheetProperties extends FlatSpec {

  private def apply(properties : SheetProperties) : XSSFSheet = {
    val sheet = new XSSFWorkbook().createSheet()
    convertSheetProperties(properties, sheet)
    sheet
  }

  private val defaultSheet : XSSFSheet = apply(SheetProperties())

  "Sheet properties conversion" should "return no auto-filter by default" in {
    //TODO Unable to test this one as auto-filter value can't be obtained
  }

  it should "return no active cell by default" in {
    assert(defaultSheet.getActiveCell == null)
  }

  it should "return B3 as active cell if explicitly specified" in {
    val model = SheetProperties(activeCell = "B3")
    val xssf = apply(model)
    assert(xssf.getActiveCell == "B3")
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
  //TODO to be continued
}
