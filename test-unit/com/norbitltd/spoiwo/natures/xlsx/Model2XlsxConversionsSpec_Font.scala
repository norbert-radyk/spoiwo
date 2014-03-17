package com.norbitltd.spoiwo.natures.xlsx

import org.apache.poi.ss.usermodel
import org.scalatest.FlatSpec
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import com.norbitltd.spoiwo.model.{Measure, Color, Font}
import Model2XlsxConversions._
import org.apache.poi.ss.usermodel.FontCharset
import Measure._

import com.norbitltd.spoiwo.model.enums.{FontFamily, Charset}

class Model2XlsxConversionsSpec_Font extends FlatSpec {

  val workbook = new XSSFWorkbook()

  "Font conversion" should "return not bold font by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    assert(!xssfDefault.getBold)
  }

  it should "return not bold font when set explicitly to not bold" in {
    val modelNotBold = Font(bold = false)
    val xssfNotBold = convertFont(modelNotBold, workbook)
    assert(!xssfNotBold.getBold)
  }

  it should "return bold font when set explicitly to bold" in {
    val modelBold = Font(bold = true)
    val xssfBold = convertFont(modelBold, workbook)
    assert(xssfBold.getBold)
  }

  it should "return ANSI charset by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    assert(FontCharset.valueOf(xssfDefault.getCharSet) === FontCharset.ANSI)
  }

  it should "return Eastern European charset when set explicitly" in {
    val modelWithCharset = Font(charSet = Charset.EastEurope)
    val xssfWithCharset = convertFont(modelWithCharset, workbook)
    assert(FontCharset.valueOf(xssfWithCharset.getCharSet) == FontCharset.EASTEUROPE)
  }

  it should "return no explicit color by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    assert(xssfDefault.getXSSFColor == null)
  }

  it should "return 'Lime' color when set explicitly" in {
    val modelWithColor = Font(color = Color.Lime)
    val xssfDefault = convertFont(modelWithColor, workbook)
    val rgbArray = xssfDefault.getXSSFColor.getRgb
    assert(rgbArray(0) == 0)
    assert(rgbArray(1) == 255.toByte)
    assert(rgbArray(2) == 0)
  }

  it should "return 'Not Applicable' font family by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    val xssfFamily = usermodel.FontFamily.valueOf(xssfDefault.getFamily)
    assert(xssfFamily == usermodel.FontFamily.NOT_APPLICABLE)
  }

  it should "return 'Roman' font family when set explicitly" in {
    val modelWithFamily = Font(family = FontFamily.Roman)
    val xssfWithFamily = convertFont(modelWithFamily, workbook)
    val xssfFamily = usermodel.FontFamily.valueOf(xssfWithFamily.getFamily)
    assert(xssfFamily == usermodel.FontFamily.ROMAN)
  }

  it should "return default height of 11 points" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    assert(11 == xssfDefault.getFontHeightInPoints)
    assert(220 == xssfDefault.getFontHeight)
  }

  it should "return height of 14 points when set explicitly" in {
    val modelWithFont = Font(height = 14 points)
    val xssfWithFont = convertFont(modelWithFont, workbook)
    assert(14 == xssfWithFont.getFontHeightInPoints)
    assert(280 == xssfWithFont.getFontHeight)
  }

  it should "return height of 160 units when set explicitly" in {
    val modelWithFont = Font(height = 160 units)
    val xssfWithFont = convertFont(modelWithFont, workbook)
    assert(8 == xssfWithFont.getFontHeightInPoints)
    assert(160 == xssfWithFont.getFontHeight)
  }


}
