package com.norbitltd.spoiwo.natures.xlsx

import org.apache.poi.ss.usermodel
import org.scalatest.{FlatSpec, Matchers}
import org.apache.poi.xssf.usermodel.{XSSFFont, XSSFWorkbook}
import com.norbitltd.spoiwo.model.{Color, Font, Height}
import Model2XlsxConversions._
import org.apache.poi.common.usermodel.fonts.FontCharset
import org.apache.poi.ss.usermodel.FontUnderline
import Height._

import scala.language.postfixOps
import com.norbitltd.spoiwo.model.enums._

class Model2XlsxConversionsForFontSpec extends FlatSpec with Matchers {

  def workbook = new XSSFWorkbook()

  def convert(f: Font): XSSFFont = convertFont(f, workbook)

  def defaultFont: XSSFFont = convert(Font())

  "Font conversion" should "return not bold font by default" in {
    defaultFont.getBold shouldBe false
  }

  it should "return not bold font when set explicitly to not bold" in {
    val model: Font = Font(bold = false)
    val xssf: XSSFFont = convert(model)
    xssf.getBold shouldBe false
  }

  it should "return bold font when set explicitly to bold" in {
    val model: Font = Font(bold = true)
    val xssf: XSSFFont = convert(model)
    xssf.getBold shouldBe true
  }

  it should "return ANSI charset by default" in {
    FontCharset.valueOf(defaultFont.getCharSet) shouldBe FontCharset.ANSI
  }

  it should "return Eastern European charset when set explicitly" in {
    val model: Font = Font(charSet = Charset.EastEurope)
    val xssf: XSSFFont = convert(model)
    FontCharset.valueOf(xssf.getCharSet) shouldBe FontCharset.EASTEUROPE
  }

  it should "return no explicit color by default" in {
    defaultFont.getXSSFColor shouldBe null
  }

  it should "return 'Lime' color when set explicitly" in {
    val model: Font = Font(color = Color.Lime)
    val xssf: XSSFFont = convert(model)
    val rgbArray = xssf.getXSSFColor.getRGB
    rgbArray(0) shouldBe 0
    rgbArray(1) shouldBe 255.toByte
    rgbArray(2) shouldBe 0
  }

  it should "return 'Not Applicable' font family by default" in {
    val xssfFamily = usermodel.FontFamily.valueOf(defaultFont.getFamily)
    xssfFamily shouldBe usermodel.FontFamily.NOT_APPLICABLE
  }

  it should "return 'Roman' font family when set explicitly" in {
    val model: Font = Font(family = FontFamily.Roman)
    val xssf: XSSFFont = convert(model)
    val xssfFamily = usermodel.FontFamily.valueOf(xssf.getFamily)
    xssfFamily shouldBe usermodel.FontFamily.ROMAN
  }

  it should "return default height of 11 points" in {
    11 shouldBe defaultFont.getFontHeightInPoints
    220 shouldBe defaultFont.getFontHeight
  }

  it should "return height of 14 points when set explicitly" in {
    val modelWithFont = Font(height = 14 points)
    val xssfWithFont = convertFont(modelWithFont, workbook)
    14 shouldBe xssfWithFont.getFontHeightInPoints
    280 shouldBe xssfWithFont.getFontHeight
  }

  it should "return height of 160 units when set explicitly" in {
    val modelWithFont = Font(height = 160 unitsOfHeight)
    val xssfWithFont = convertFont(modelWithFont, workbook)
    8 shouldBe xssfWithFont.getFontHeightInPoints
    160 shouldBe xssfWithFont.getFontHeight
  }

  it should "return not italic font by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    xssfDefault.getItalic shouldBe false
  }

  it should "return not italic font when set explicitly to not italic" in {
    val modelNotItalic = Font(italic = false)
    val xssfNotItalic = convertFont(modelNotItalic, workbook)
    xssfNotItalic.getItalic shouldBe false
  }

  it should "return italic font when set explicitly to italic" in {
    val modelItalic = Font(italic = true)
    val xssfItalic = convertFont(modelItalic, workbook)
    xssfItalic.getItalic shouldBe true
  }

  it should "return no font scheme by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    xssfDefault.getScheme shouldBe usermodel.FontScheme.NONE
  }

  it should "return 'Major' font scheme when set explicitly" in {
    val modelWithScheme = Font(scheme = FontScheme.Major)
    val xssfWithScheme = convertFont(modelWithScheme, workbook)
    xssfWithScheme.getScheme shouldBe usermodel.FontScheme.MAJOR
  }

  it should "return 'Calibri' font name by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    xssfDefault.getFontName shouldBe "Calibri"
  }

  it should "return 'Arial' font name when set explicitly" in {
    val modelWithFontName = Font(fontName = "Arial")
    val xssfWithFontName = convertFont(modelWithFontName, workbook)
    xssfWithFontName.getFontName shouldBe "Arial"
  }

  it should "return not strikeout font by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    xssfDefault.getStrikeout shouldBe false
  }

  it should "return not strikeout font when set explicitly to not strikeout" in {
    val modelNotStrikeout = Font(strikeout = false)
    val xssfNotStrikeout = convertFont(modelNotStrikeout, workbook)
    xssfNotStrikeout.getStrikeout shouldBe false
  }

  it should "return strikeout font when set explicitly to strikeout" in {
    val modelStrikeout = Font(strikeout = true)
    val xssfStrikeout = convertFont(modelStrikeout, workbook)
    xssfStrikeout.getStrikeout shouldBe true
  }

  it should "return no type offset by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    xssfDefault.getTypeOffset shouldBe usermodel.Font.SS_NONE
  }

  it should "return 'Subscript' type offset when set explicitly" in {
    val modelWithTypeOffset = Font(typeOffset = TypeOffset.Subscript)
    val xssfWithTypeOffset = convertFont(modelWithTypeOffset, workbook)
    xssfWithTypeOffset.getTypeOffset shouldBe usermodel.Font.SS_SUB
  }

  it should "return no underline by default" in {
    val modelDefault = Font()
    val xssfDefault = convertFont(modelDefault, workbook)
    xssfDefault.getUnderline shouldBe FontUnderline.NONE.getByteValue
  }

  it should "return double underline when set explicitly" in {
    val modelWithUnderline = Font(underline = Underline.Double)
    val xssfWithUnderline = convertFont(modelWithUnderline, workbook)
    xssfWithUnderline.getUnderline shouldBe FontUnderline.DOUBLE.getByteValue
  }

  //Next 4 test cases will test enhanced boolean conversion
  it should "return no bold, no italic by default" in {
    defaultFont.getBold shouldBe false
    defaultFont.getItalic shouldBe false
  }

  it should "return bold, but no italic when set" in {
    val xssf = convert(Font(bold = true))
    xssf.getBold shouldBe true
    xssf.getItalic shouldBe false
  }

  it should "return no bold, but italic when set" in {
    val xssf = convert(Font(italic = true))
    xssf.getBold shouldBe false
    xssf.getItalic shouldBe true
  }

  it should "return bold and italic when set" in {
    val xssf = convert(Font(italic = true, bold = true))
    xssf.getBold shouldBe true
    xssf.getItalic shouldBe true
  }
}
