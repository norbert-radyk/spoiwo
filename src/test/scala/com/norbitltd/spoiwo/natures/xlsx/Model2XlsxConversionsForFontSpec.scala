package com.norbitltd.spoiwo.natures.xlsx

import org.apache.poi.ss.usermodel
import org.scalatest.FlatSpec
import org.apache.poi.xssf.usermodel.{XSSFFont, XSSFWorkbook}
import com.norbitltd.spoiwo.model.{Height, Color, Font}
import Model2XlsxConversions._
import org.apache.poi.ss.usermodel.{FontUnderline, FontCharset}
import Height._

import com.norbitltd.spoiwo.model.enums._

class Model2XlsxConversionsForFontSpec extends FlatSpec {

  def workbook = new XSSFWorkbook()

  def convert(f : Font) : XSSFFont = convertFont(f, workbook)

  def defaultFont : XSSFFont = convert(Font())

  "Font conversion" should "return not bold font by default" in {
    assert(!defaultFont.getBold)
  }

  it should "return not bold font when set explicitly to not bold" in {
    val model : Font = Font(bold = false)
    val xssf : XSSFFont = convert(model)
    assert(!xssf.getBold)
  }

  it should "return bold font when set explicitly to bold" in {
    val model : Font = Font(bold = true)
    val xssf : XSSFFont = convert(model)
    assert(xssf.getBold)
  }

  it should "return ANSI charset by default" in {
    assert(FontCharset.valueOf(defaultFont.getCharSet) === FontCharset.ANSI)
  }

  it should "return Eastern European charset when set explicitly" in {
    val model : Font = Font(charSet = Charset.EastEurope)
    val xssf : XSSFFont = convert(model)
    assert(FontCharset.valueOf(xssf.getCharSet) == FontCharset.EASTEUROPE)
  }

  it should "return no explicit color by default" in {
    assert(defaultFont.getXSSFColor == null)
  }

  it should "return 'Lime' color when set explicitly" in {
    val model : Font = Font(color = Color.Lime)
    val xssf : XSSFFont = convert(model)
    val rgbArray = xssf.getXSSFColor.getRgb
    assert(rgbArray(0) == 0)
    assert(rgbArray(1) == 255.toByte)
    assert(rgbArray(2) == 0)
  }

  it should "return 'Not Applicable' font family by default" in {
    val xssfFamily = usermodel.FontFamily.valueOf(defaultFont.getFamily)
    assert(xssfFamily == usermodel.FontFamily.NOT_APPLICABLE)
  }

  it should "return 'Roman' font family when set explicitly" in {
    val model : Font = Font(family = FontFamily.Roman)
    val xssf : XSSFFont = convert(model)
    val xssfFamily = usermodel.FontFamily.valueOf(xssf.getFamily)
    assert(xssfFamily == usermodel.FontFamily.ROMAN)
  }

  it should "return default height of 11 points" in {
    assert(11 == defaultFont.getFontHeightInPoints)
    assert(220 == defaultFont.getFontHeight)
  }

  it should "return height of 14 points when set explicitly" in {
    val modelWithFont = Font(height = 14 points)
    val xssfWithFont = convertFont(modelWithFont, workbook)
    assert(14 == xssfWithFont.getFontHeightInPoints)
    assert(280 == xssfWithFont.getFontHeight)
  }

  it should "return height of 160 units when set explicitly" in {
    val modelWithFont = Font(height = 160 unitsOfHeight)
    val xssfWithFont = convertFont(modelWithFont, workbook)
    assert(8 == xssfWithFont.getFontHeightInPoints)
    assert(160 == xssfWithFont.getFontHeight)
  }

  it should "return not italic font by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    assert(!xssfDefault.getItalic)
  }

  it should "return not italic font when set explicitly to not italic" in {
    val modelNotItalic = Font(italic = false)
    val xssfNotItalic = convertFont(modelNotItalic, workbook)
    assert(!xssfNotItalic.getItalic)
  }

  it should "return italic font when set explicitly to italic" in {
    val modelItalic = Font(italic = true)
    val xssfItalic = convertFont(modelItalic, workbook)
    assert(xssfItalic.getItalic)
  }

  it should "return no font scheme by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    assert(xssfDefault.getScheme == usermodel.FontScheme.NONE)
  }

  it should "return 'Major' font scheme when set explicitly" in {
    val modelWithScheme = Font(scheme = FontScheme.Major)
    val xssfWithScheme = convertFont(modelWithScheme, workbook)
    assert(xssfWithScheme.getScheme == usermodel.FontScheme.MAJOR)
  }

  it should "return 'Calibri' font name by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    assert(xssfDefault.getFontName == "Calibri")
  }

  it should "return 'Arial' font name when set explicitly" in {
    val modelWithFontName = Font(fontName = "Arial")
    val xssfWithFontName = convertFont(modelWithFontName, workbook)
    assert(xssfWithFontName.getFontName == "Arial")
  }

  it should "return not strikeout font by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    assert(!xssfDefault.getStrikeout)
  }

  it should "return not strikeout font when set explicitly to not strikeout" in {
    val modelNotStrikeout = Font(strikeout = false)
    val xssfNotStrikeout = convertFont(modelNotStrikeout, workbook)
    assert(!xssfNotStrikeout.getStrikeout)
  }

  it should "return strikeout font when set explicitly to strikeout" in {
    val modelStrikeout = Font(strikeout = true)
    val xssfStrikeout = convertFont(modelStrikeout, workbook)
    assert(xssfStrikeout.getStrikeout)
  }

  it should "return no type offset by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    assert(xssfDefault.getTypeOffset == usermodel.Font.SS_NONE)
  }

  it should "return 'Subscript' type offset when set explicitly" in {
    val modelWithTypeOffset = Font(typeOffset = TypeOffset.Subscript)
    val xssfWithTypeOffset = convertFont(modelWithTypeOffset, workbook)
    assert(xssfWithTypeOffset.getTypeOffset == usermodel.Font.SS_SUB)
  }

  it should "return no underline by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    assert(xssfDefault.getUnderline == FontUnderline.NONE.getByteValue)
  }

  it should "return double underline when set explicitly" in {
    val modelWithUnderline = Font(underline = Underline.Double)
    val xssfWithUnderline = convertFont(modelWithUnderline, workbook)
    assert(xssfWithUnderline.getUnderline == FontUnderline.DOUBLE.getByteValue)
  }

  //Next 4 test cases will test enhanced boolean conversion
  it should "return no bold, no italic by default" in {
    assert(!defaultFont.getBold)
    assert(!defaultFont.getItalic)
  }

  it should "return bold, but no italic when set" in {
    val xssf = convert(Font(bold = true))
    assert(xssf.getBold)
    assert(!xssf.getItalic)
  }

  it should "return no bold, but italic when set" in {
    val xssf = convert(Font(italic = true))
    assert(!xssf.getBold)
    assert(xssf.getItalic)
  }

  it should "return bold and italic when set" in {
    val xssf = convert(Font(italic = true, bold = true))
    assert(xssf.getBold)
    assert(xssf.getItalic)
  }
}