package com.norbitltd.spoiwo.natures.xlsx

import org.scalatest.FlatSpec
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import com.norbitltd.spoiwo.model.{Color, Font}
import Model2XlsxConversions._
import org.apache.poi.ss.usermodel.FontCharset
import com.norbitltd.spoiwo.model.enums.Charset

class Model2XlsxConversionsSpec_Font extends FlatSpec {

  val workbook = new XSSFWorkbook()

  "Font conversion" should "return not bold font by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    assert(!xssfDefault.getBold)
  }

  it should "return not bold font when explicitly set to not bold" in {
    val modelNotBold = Font(bold = false)
    val xssfNotBold = convertFont(modelNotBold, workbook)
    assert(!xssfNotBold.getBold)
  }

  it should "return bold font when explicitly set to bold" in {
    val modelBold = Font(bold = true)
    val xssfBold = convertFont(modelBold, workbook)
    assert(xssfBold.getBold)
  }

  it should "return ANSI charset by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    assert(FontCharset.valueOf(xssfDefault.getCharSet) === FontCharset.ANSI)
  }

  it should "return Eastern European charset when explicitly specified" in {
    val modelWithCharset = Font(charSet = Charset.EastEurope)
    val xssfWithCharset = convertFont(modelWithCharset, workbook)
    assert(FontCharset.valueOf(xssfWithCharset.getCharSet) == FontCharset.EASTEUROPE)
  }

  it should "return no explicit color by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    assert(xssfDefault.getXSSFColor == null)
  }

  it should "return lime color when explicitly specified" in {
    val modelWithColor = Font(color = Color.Lime)
    val xssfDefault = convertFont(modelWithColor, workbook)
    val rgbArray = xssfDefault.getXSSFColor.getRgb
    assert(rgbArray(0) == 0)
    assert(rgbArray(1) == 255.toByte)
    assert(rgbArray(2) == 0)
  }


}
