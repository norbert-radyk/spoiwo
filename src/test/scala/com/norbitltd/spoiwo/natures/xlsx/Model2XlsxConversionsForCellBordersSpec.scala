package com.norbitltd.spoiwo.natures.xlsx

import org.scalatest.FlatSpec
import Model2XlsxConversions.convertCellBorders
import org.apache.poi.ss.usermodel
import org.apache.poi.xssf.usermodel.{XSSFCellStyle, XSSFWorkbook}
import com.norbitltd.spoiwo.model.{Color, CellBorders}
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder
import com.norbitltd.spoiwo.model.enums.CellBorderStyle

class Model2XlsxConversionsForCellBordersSpec extends FlatSpec{

  val cellStyle : XSSFCellStyle = new XSSFWorkbook().createCellStyle()

  def convert(cellBorders : CellBorders) : XSSFCellStyle = {
    val style : XSSFCellStyle = cellStyle
    convertCellBorders(cellBorders, style)
    style
  }

  val default = convert(CellBorders())

  "Cell borders conversion" should "return no borders by default" in {
    assert(default.getBorderBottomEnum == usermodel.BorderStyle.NONE)
    assert(default.getBorderColor(XSSFCellBorder.BorderSide.BOTTOM) == null)
    assert(default.getBorderTopEnum == usermodel.BorderStyle.NONE)
    assert(default.getBorderColor(XSSFCellBorder.BorderSide.TOP) == null)
    assert(default.getBorderLeftEnum == usermodel.BorderStyle.NONE)
    assert(default.getBorderColor(XSSFCellBorder.BorderSide.LEFT) == null)
    assert(default.getBorderRightEnum == usermodel.BorderStyle.NONE)
    assert(default.getBorderColor(XSSFCellBorder.BorderSide.RIGHT) == null)
  }

  it should "return user defined bottom border" in {
    val model = CellBorders(bottomStyle = CellBorderStyle.Dashed, bottomColor = Color.Blue)
    val xlsx = convert(model)

    assert(xlsx.getBorderBottomEnum == usermodel.BorderStyle.DASHED)
    assert(xlsx.getBorderColor(XSSFCellBorder.BorderSide.BOTTOM).getRGB.toList ==
      Array[Byte](0, 0, 255.toByte).toList)
  }

  it should "return user defined top border" in {
    val model = CellBorders(topStyle = CellBorderStyle.MediumDashDotDot, topColor = Color.Olive)
    val xlsx = convert(model)

    assert(xlsx.getBorderTopEnum == usermodel.BorderStyle.MEDIUM_DASH_DOT_DOTC)
    assert(xlsx.getBorderColor(XSSFCellBorder.BorderSide.TOP).getRGB.toList ==
      Array[Byte](128.toByte, 128.toByte, 0.toByte).toList)
  }

  it should "return user defined left border" in {
    val model = CellBorders(leftStyle = CellBorderStyle.Dotted, leftColor = Color.Blue)
    val xlsx = convert(model)

    assert(xlsx.getBorderLeftEnum == usermodel.BorderStyle.DOTTED)
    assert(xlsx.getBorderColor(XSSFCellBorder.BorderSide.LEFT).getRGB.toList ==
      Array[Byte](0, 0, 255.toByte).toList)
  }

  it should "return user defined right border" in {
    val model = CellBorders(rightStyle = CellBorderStyle.Hair, rightColor = Color.Blue)
    val xlsx = convert(model)

    assert(xlsx.getBorderRightEnum == usermodel.BorderStyle.HAIR)
    assert(xlsx.getBorderColor(XSSFCellBorder.BorderSide.RIGHT).getRGB.toList ==
      Array[Byte](0, 0, 255.toByte).toList)
  }

}
