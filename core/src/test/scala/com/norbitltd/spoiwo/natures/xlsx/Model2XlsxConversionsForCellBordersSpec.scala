package com.norbitltd.spoiwo.natures.xlsx

import org.scalatest.{FlatSpec, Matchers}
import Model2XlsxConversions.convertCellBorders
import org.apache.poi.ss.usermodel
import org.apache.poi.xssf.usermodel.{XSSFCellStyle, XSSFWorkbook}
import com.norbitltd.spoiwo.model.{CellBorders, Color}
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder
import com.norbitltd.spoiwo.model.enums.CellBorderStyle

class Model2XlsxConversionsForCellBordersSpec extends FlatSpec with Matchers {

  val cellStyle: XSSFCellStyle = new XSSFWorkbook().createCellStyle()

  def convert(cellBorders: CellBorders): XSSFCellStyle = {
    val style: XSSFCellStyle = cellStyle
    convertCellBorders(cellBorders, style)
    style
  }

  val default: XSSFCellStyle = convert(CellBorders())

  "Cell borders conversion" should "return no borders by default" in {
    default.getBorderBottomEnum shouldBe usermodel.BorderStyle.NONE
    default.getBorderColor(XSSFCellBorder.BorderSide.BOTTOM) shouldBe null
    default.getBorderTopEnum shouldBe usermodel.BorderStyle.NONE
    default.getBorderColor(XSSFCellBorder.BorderSide.TOP) shouldBe null
    default.getBorderLeftEnum shouldBe usermodel.BorderStyle.NONE
    default.getBorderColor(XSSFCellBorder.BorderSide.LEFT) shouldBe null
    default.getBorderRightEnum shouldBe usermodel.BorderStyle.NONE
    default.getBorderColor(XSSFCellBorder.BorderSide.RIGHT) shouldBe null
  }

  it should "return user defined bottom border" in {
    val model = CellBorders(bottomStyle = CellBorderStyle.Dashed, bottomColor = Color.Blue)
    val xlsx = convert(model)

    xlsx.getBorderBottomEnum shouldBe usermodel.BorderStyle.DASHED
    xlsx.getBorderColor(XSSFCellBorder.BorderSide.BOTTOM).getRGB.toList shouldBe
      Array[Byte](0, 0, 255.toByte).toList
  }

  it should "return user defined top border" in {
    val model = CellBorders(topStyle = CellBorderStyle.DashDot, topColor = Color.Olive)
    val xlsx = convert(model)

    xlsx.getBorderTopEnum shouldBe usermodel.BorderStyle.DASH_DOT
    xlsx.getBorderColor(XSSFCellBorder.BorderSide.TOP).getRGB.toList shouldBe
      Array[Byte](128.toByte, 128.toByte, 0.toByte).toList
  }

  it should "return user defined left border" in {
    val model = CellBorders(leftStyle = CellBorderStyle.Dotted, leftColor = Color.Blue)
    val xlsx = convert(model)

    xlsx.getBorderLeftEnum shouldBe usermodel.BorderStyle.DOTTED
    xlsx.getBorderColor(XSSFCellBorder.BorderSide.LEFT).getRGB.toList shouldBe
      Array[Byte](0, 0, 255.toByte).toList
  }

  it should "return user defined right border" in {
    val model = CellBorders(rightStyle = CellBorderStyle.Hair, rightColor = Color.Blue)
    val xlsx = convert(model)

    xlsx.getBorderRightEnum shouldBe usermodel.BorderStyle.HAIR
    xlsx.getBorderColor(XSSFCellBorder.BorderSide.RIGHT).getRGB.toList shouldBe
      Array[Byte](0, 0, 255.toByte).toList
  }

}
