package com.norbitltd.spoiwo.natures.xlsx

import org.scalatest.{FlatSpec, Matchers}
import org.apache.poi.xssf.usermodel.{XSSFCellStyle, XSSFFont, XSSFWorkbook}
import com.norbitltd.spoiwo.model._
import Model2XlsxConversions._
import org.apache.poi.ss.usermodel.{BorderStyle, FillPatternType, ReadingOrder, HorizontalAlignment, VerticalAlignment}
import Height._
import com.norbitltd.spoiwo.model.enums.{CellBorderStyle, CellFill, CellReadingOrder, CellHorizontalAlignment, CellVerticalAlignment}

import scala.language.postfixOps

class Model2XlsxConversionsForCellStyleSpec extends FlatSpec with Matchers {

  val workbook = new XSSFWorkbook()

  "Cell style conversion" should "return no borders by default" in {
    val modelDefault = CellStyle()
    val xssfDefault = convertCellStyle(modelDefault, workbook)
    xssfDefault.getBorderBottomEnum shouldBe BorderStyle.NONE
    xssfDefault.getBorderLeftEnum shouldBe BorderStyle.NONE
    xssfDefault.getBorderRightEnum shouldBe BorderStyle.NONE
    xssfDefault.getBorderTopEnum shouldBe BorderStyle.NONE
    xssfDefault.getBottomBorderXSSFColor shouldBe null
    xssfDefault.getTopBorderXSSFColor shouldBe null
    xssfDefault.getLeftBorderXSSFColor shouldBe null
    xssfDefault.getRightBorderXSSFColor shouldBe null
  }

  it should "return explicitly set borders" in {
    val modelBorders = CellBorders(leftStyle = CellBorderStyle.Double,
                                   leftColor = Color.Blue,
                                   rightStyle = CellBorderStyle.Thick,
                                   rightColor = Color.Red)
    val modelWithBorders = CellStyle(borders = modelBorders)
    val xssfWithBorders = convertCellStyle(modelWithBorders, workbook)

    xssfWithBorders.getBorderLeftEnum shouldBe BorderStyle.DOUBLE
    xssfWithBorders.getBorderRightEnum shouldBe BorderStyle.THICK

    val leftColorRGB = xssfWithBorders.getLeftBorderXSSFColor.getRGB
    leftColorRGB.toList shouldBe List(0, 0, 255.toByte)
    val rightColorRGB = xssfWithBorders.getRightBorderXSSFColor.getRGB
    rightColorRGB.toList shouldBe List(255.toByte, 0, 0)
  }

  it should "return 'General' data format by default" in {
    val modelDefault = CellStyle()
    val xssfDefault = convertCellStyle(modelDefault, workbook)
    xssfDefault.getDataFormatString shouldBe "General"
  }

  it should "return 'yyyy-MM-dd' data format when set explicitly" in {
    val modelWithFormat = CellStyle(dataFormat = CellDataFormat("yyyy-MM-dd"))
    val xssfWithFormat = convertCellStyle(modelWithFormat, workbook)
    xssfWithFormat.getDataFormatString shouldBe "yyyy-MM-dd"
  }

  it should "return 11pt Calibri font by default" in {
    val modelDefault = CellStyle()
    val xssfDefault = convertCellStyle(modelDefault, workbook)
    val xssfFont = xssfDefault.getFont
    xssfFont.getFontName shouldBe "Calibri"
    xssfFont.getFontHeightInPoints shouldBe 11
  }

  it should "return 14pt Arial when set explicitly" in {
    val modelWithFont: CellStyle = CellStyle(font = Font(fontName = "Arial", height = 14 points))
    val xssfWithFont: XSSFCellStyle = convertCellStyle(modelWithFont, workbook)
    val xssfFont: XSSFFont = xssfWithFont.getFont
    xssfFont.getFontName shouldBe "Arial"
    xssfFont.getFontHeightInPoints shouldBe 14
  }

  it should "return no fill pattern by default" in {
    val modelDefault = CellStyle()
    val xssfDefault = convertCellStyle(modelDefault, workbook)
    xssfDefault.getFillPatternEnum shouldBe FillPatternType.NO_FILL
  }

  it should "return 'Solid' fill pattern when explicitly set" in {
    val modelWithFillPattern: CellStyle = CellStyle(fillPattern = CellFill.Solid)
    val xssfWithFillPattern: XSSFCellStyle = convertCellStyle(modelWithFillPattern, workbook)
    xssfWithFillPattern.getFillPatternEnum shouldBe FillPatternType.SOLID_FOREGROUND
  }

  it should "return no background or foreground color by default" in {
    val modelDefault: CellStyle = CellStyle()
    val xssfDefault: XSSFCellStyle = convertCellStyle(modelDefault, workbook)
    xssfDefault.getFillForegroundXSSFColor shouldBe null
    xssfDefault.getFillBackgroundXSSFColor shouldBe null
  }

  it should "return red background and blue foreground color when explicitly set" in {
    val modelWithFillColors: CellStyle = CellStyle(fillBackgroundColor = Color.Red, fillForegroundColor = Color.Blue)
    val xssfDefault: XSSFCellStyle = convertCellStyle(modelWithFillColors, workbook)
    xssfDefault.getFillBackgroundXSSFColor.getRGB.toList shouldBe List(255.toByte, 0, 0)
    xssfDefault.getFillForegroundXSSFColor.getRGB.toList shouldBe List(0, 0, 255.toByte)
  }

  it should "return 'Context' reading order by default" in {
    val modelDefault: CellStyle = CellStyle()
    val xssfDefault: XSSFCellStyle = convertCellStyle(modelDefault, workbook)
    xssfDefault.getReadingOrder shouldBe ReadingOrder.CONTEXT
  }

  it should "return 'Right to Left' reading order when explicitly set" in {
    val modelWithHA: CellStyle = CellStyle(readingOrder = CellReadingOrder.RightToLeft)
    val xssfWithHA: XSSFCellStyle = convertCellStyle(modelWithHA, workbook)
    xssfWithHA.getReadingOrder shouldBe ReadingOrder.RIGHT_TO_LEFT
  }

  it should "return 'General' horizontal alignment by default" in {
    val modelDefault: CellStyle = CellStyle()
    val xssfDefault: XSSFCellStyle = convertCellStyle(modelDefault, workbook)
    xssfDefault.getAlignmentEnum shouldBe HorizontalAlignment.GENERAL
  }

  it should "return 'Right' horizontal alignment when explicitly set" in {
    val modelWithHA: CellStyle = CellStyle(horizontalAlignment = CellHorizontalAlignment.Right)
    val xssfWithHA: XSSFCellStyle = convertCellStyle(modelWithHA, workbook)
    xssfWithHA.getAlignmentEnum shouldBe HorizontalAlignment.RIGHT
  }

  it should "return 'Bottom' vertical alignment by default" in {
    val modelDefault: CellStyle = CellStyle()
    val xssfDefault: XSSFCellStyle = convertCellStyle(modelDefault, workbook)
    xssfDefault.getVerticalAlignmentEnum shouldBe VerticalAlignment.BOTTOM
  }

  it should "return 'Top' vertical alignment when explicitly set" in {
    val modelWithVA: CellStyle = CellStyle(verticalAlignment = CellVerticalAlignment.Top)
    val xssfWithVA: XSSFCellStyle = convertCellStyle(modelWithVA, workbook)
    xssfWithVA.getVerticalAlignmentEnum shouldBe VerticalAlignment.TOP
  }

  it should "return not hidden cell style by default" in {
    val modelDefault: CellStyle = CellStyle()
    val xssfDefault: XSSFCellStyle = convertCellStyle(modelDefault, workbook)
    xssfDefault.getHidden shouldBe false
  }

  it should "return not hidden cell style when explicitly set to false" in {
    val model: CellStyle = CellStyle(hidden = false)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    xssf.getHidden shouldBe false
  }

  it should "return hidden cell style when explicitly set to true" in {
    val model: CellStyle = CellStyle(hidden = true)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    xssf.getHidden shouldBe true
  }

  it should "return no indention by default" in {
    val model: CellStyle = CellStyle()
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    xssf.getIndention shouldBe 0
  }

  it should "return indention of 4 spaces when explicitly set" in {
    val model: CellStyle = CellStyle(indention = 4)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    xssf.getIndention shouldBe 4
  }

  it should "return locked cell by default" in {
    val model: CellStyle = CellStyle()
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    xssf.getLocked shouldBe true
  }

  it should "return unlocked cell when explicitly set to false" in {
    val model: CellStyle = CellStyle(locked = false)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    xssf.getLocked shouldBe false
  }

  it should "return locked cell when explicitly set to true" in {
    val model: CellStyle = CellStyle(locked = true)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    xssf.getLocked shouldBe true
  }

  it should "return no rotation by default" in {
    val model: CellStyle = CellStyle()
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    xssf.getRotation shouldBe 0
  }

  it should "return 90 degree rotation if explicitly set" in {
    val model: CellStyle = CellStyle(rotation = 90)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    xssf.getRotation shouldBe 90
  }

  it should "return unwrapped text by default" in {
    val model: CellStyle = CellStyle()
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    xssf.getWrapText shouldBe false
  }

  it should "return unwrapped text when explicitly set to false" in {
    val model: CellStyle = CellStyle(wrapText = false)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    xssf.getWrapText shouldBe false
  }

  it should "return wrapped text when explicitly set to true" in {
    val model: CellStyle = CellStyle(wrapText = true)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    xssf.getWrapText shouldBe true
  }

}
