package com.norbitltd.spoiwo.natures.xlsx

import org.scalatest.FlatSpec
import org.apache.poi.xssf.usermodel.{XSSFFont, XSSFCellStyle, XSSFWorkbook}
import com.norbitltd.spoiwo.model._
import Model2XlsxConversions._
import org.apache.poi.ss.usermodel.{VerticalAlignment, HorizontalAlignment, FillPatternType, BorderStyle}
import Height._
import com.norbitltd.spoiwo.model.enums.{CellBorderStyle, CellVerticalAlignment, CellHorizontalAlignment, CellFill}
import scala.language.postfixOps

class Model2XlsxConversionsForCellStyleSpec extends FlatSpec {

  val workbook = new XSSFWorkbook()

  "Cell style conversion" should "return no borders by default" in {
    val modelDefault = CellStyle()
    val xssfDefault = convertCellStyle(modelDefault, workbook)
    assert(xssfDefault.getBorderBottomEnum == BorderStyle.NONE)
    assert(xssfDefault.getBorderLeftEnum == BorderStyle.NONE)
    assert(xssfDefault.getBorderRightEnum == BorderStyle.NONE)
    assert(xssfDefault.getBorderTopEnum == BorderStyle.NONE)
    assert(xssfDefault.getBottomBorderXSSFColor == null)
    assert(xssfDefault.getTopBorderXSSFColor == null)
    assert(xssfDefault.getLeftBorderXSSFColor == null)
    assert(xssfDefault.getRightBorderXSSFColor == null)
  }

  it should "return explicitly set borders" in {
    val modelBorders = CellBorders(
      leftStyle = CellBorderStyle.Double,
      leftColor = Color.Blue,
      rightStyle = CellBorderStyle.Thick,
      rightColor = Color.Red)
    val modelWithBorders = CellStyle(borders = modelBorders)
    val xssfWithBorders = convertCellStyle(modelWithBorders, workbook)

    assert(xssfWithBorders.getBorderLeftEnum == BorderStyle.DOUBLE)
    assert(xssfWithBorders.getBorderRightEnum == BorderStyle.THICK)

    val leftColorRGB = xssfWithBorders.getLeftBorderXSSFColor.getRGB
    assert(leftColorRGB.toList == List(0, 0, 255.toByte))
    val rightColorRGB = xssfWithBorders.getRightBorderXSSFColor.getRGB
    assert(rightColorRGB.toList == List(255.toByte, 0, 0))
  }

  it should "return 'General' data format by default" in {
    val modelDefault = CellStyle()
    val xssfDefault = convertCellStyle(modelDefault, workbook)
    assert(xssfDefault.getDataFormatString == "General")
  }

  it should "return 'yyyy-MM-dd' data format when set explicitly" in {
    val modelWithFormat = CellStyle(dataFormat = CellDataFormat("yyyy-MM-dd"))
    val xssfWithFormat = convertCellStyle(modelWithFormat, workbook)
    assert(xssfWithFormat.getDataFormatString == "yyyy-MM-dd")
  }

  it should "return 11pt Calibri font by default" in {
    val modelDefault = CellStyle()
    val xssfDefault = convertCellStyle(modelDefault, workbook)
    val xssfFont = xssfDefault.getFont
    assert(xssfFont.getFontName == "Calibri")
    assert(xssfFont.getFontHeightInPoints == 11)
  }

  it should "return 14pt Arial when set explicitly" in {
    val modelWithFont : CellStyle = CellStyle(font = Font(fontName = "Arial", height = 14 points))
    val xssfWithFont : XSSFCellStyle = convertCellStyle(modelWithFont, workbook)
    val xssfFont : XSSFFont = xssfWithFont.getFont
    assert(xssfFont.getFontName == "Arial")
    assert(xssfFont.getFontHeightInPoints == 14)
  }

  it should "return no fill pattern by default" in {
    val modelDefault = CellStyle()
    val xssfDefault = convertCellStyle(modelDefault, workbook)
    assert(xssfDefault.getFillPatternEnum == FillPatternType.NO_FILL)
  }

  it should "return 'Solid' fill pattern when explicitly set" in {
    val modelWithFillPattern : CellStyle = CellStyle(fillPattern = CellFill.Solid)
    val xssfWithFillPattern : XSSFCellStyle = convertCellStyle(modelWithFillPattern, workbook)
    assert(xssfWithFillPattern.getFillPatternEnum == FillPatternType.SOLID_FOREGROUND)
  }

  it should "return no background or foreground color by default" in {
    val modelDefault : CellStyle = CellStyle()
    val xssfDefault : XSSFCellStyle = convertCellStyle(modelDefault, workbook)
    assert(xssfDefault.getFillForegroundXSSFColor == null)
    assert(xssfDefault.getFillBackgroundXSSFColor == null)
  }

  it should "return red background and blue foreground color when explicitly set" in {
    val modelWithFillColors : CellStyle = CellStyle(fillBackgroundColor = Color.Red, fillForegroundColor = Color.Blue)
    val xssfDefault : XSSFCellStyle = convertCellStyle(modelWithFillColors, workbook)
    assert(xssfDefault.getFillBackgroundXSSFColor.getRGB.toList == List(255.toByte, 0, 0))
    assert(xssfDefault.getFillForegroundXSSFColor.getRGB.toList == List(0, 0, 255.toByte))
  }

  it should "return 'General' horizontal alignment by default" in {
    val modelDefault : CellStyle = CellStyle()
    val xssfDefault : XSSFCellStyle = convertCellStyle(modelDefault, workbook)
    assert(xssfDefault.getAlignmentEnum == HorizontalAlignment.GENERAL)
  }

  it should "return 'Right' horizontal alignment when explicitly set" in {
    val modelWithHA : CellStyle = CellStyle(horizontalAlignment = CellHorizontalAlignment.Right)
    val xssfWithHA : XSSFCellStyle = convertCellStyle(modelWithHA, workbook)
    assert(xssfWithHA.getAlignmentEnum == HorizontalAlignment.RIGHT)
  }

  it should "return 'Bottom' vertical alignment by default" in {
    val modelDefault : CellStyle = CellStyle()
    val xssfDefault : XSSFCellStyle = convertCellStyle(modelDefault, workbook)
    assert(xssfDefault.getVerticalAlignmentEnum == VerticalAlignment.BOTTOM)
  }

  it should "return 'Top' vertical alignment when explicitly set" in {
    val modelWithVA : CellStyle = CellStyle(verticalAlignment = CellVerticalAlignment.Top)
    val xssfWithVA : XSSFCellStyle = convertCellStyle(modelWithVA, workbook)
    assert(xssfWithVA.getVerticalAlignmentEnum == VerticalAlignment.TOP)
  }

  it should "return not hidden cell style by default" in {
    val modelDefault : CellStyle = CellStyle()
    val xssfDefault: XSSFCellStyle = convertCellStyle(modelDefault, workbook)
    assert(!xssfDefault.getHidden)
  }

  it should "return not hidden cell style when explicitly set to false" in {
    val model : CellStyle = CellStyle(hidden = false)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    assert(!xssf.getHidden)
  }

  it should "return hidden cell style when explicitly set to true" in {
    val model : CellStyle = CellStyle(hidden = true)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    assert(xssf.getHidden)
  }

  it should "return no indention by default" in {
    val model : CellStyle = CellStyle()
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    assert(xssf.getIndention == 0)
  }

  it should "return indention of 4 spaces when explicitly set" in {
    val model : CellStyle = CellStyle(indention = 4)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    assert(xssf.getIndention == 4)
  }

  it should "return locked cell by default" in {
    val model : CellStyle = CellStyle()
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    assert(xssf.getLocked)
  }

  it should "return unlocked cell when explicitly set to false" in {
    val model : CellStyle = CellStyle(locked = false)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    assert(!xssf.getLocked)
  }

  it should "return locked cell when explicitly set to true" in {
    val model : CellStyle = CellStyle(locked = true)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    assert(xssf.getLocked)
  }

  it should "return no rotation by default" in {
    val model : CellStyle = CellStyle()
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    assert(xssf.getRotation == 0)
  }

  it should "return 90 degree rotation if explicitly set" in {
    val model : CellStyle = CellStyle(rotation = 90)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    assert(xssf.getRotation == 90)
  }

  it should "return unwrapped text by default" in {
    val model : CellStyle = CellStyle()
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    assert(!xssf.getWrapText)
  }

  it should "return unwrapped text when explicitly set to false" in {
    val model : CellStyle = CellStyle(wrapText = false)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    assert(!xssf.getWrapText)
  }

  it should "return wrapped text when explicitly set to true" in {
    val model : CellStyle = CellStyle(wrapText = true)
    val xssf: XSSFCellStyle = convertCellStyle(model, workbook)
    assert(xssf.getWrapText)
  }

}
