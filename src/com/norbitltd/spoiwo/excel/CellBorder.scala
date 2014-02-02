package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.xssf.usermodel.XSSFCellStyle

object CellBorder extends Factory {

  private lazy val defaultLeftStyle = defaultPOICellStyle.getBorderLeftEnum
  private lazy val defaultLeftColor = Color(defaultPOICellStyle.getLeftBorderXSSFColor)
  private lazy val defaultTopStyle = defaultPOICellStyle.getBorderTopEnum
  private lazy val defaultTopColor = Color(defaultPOICellStyle.getTopBorderXSSFColor)
  private lazy val defaultRightStyle = defaultPOICellStyle.getBorderRightEnum
  private lazy val defaultRightColor = Color(defaultPOICellStyle.getRightBorderXSSFColor)
  private lazy val defaultBottomStyle = defaultPOICellStyle.getBorderBottomEnum
  private lazy val defaultBottomColor = Color(defaultPOICellStyle.getBottomBorderXSSFColor)

  val Default = CellBorder()

  def apply(leftStyle: BorderStyle = defaultLeftStyle, leftColor: Color = defaultLeftColor,
            topStyle: BorderStyle = defaultTopStyle, topColor: Color = defaultTopColor,
            rightStyle: BorderStyle = defaultRightStyle, rightColor: Color = defaultRightColor,
            bottomStyle: BorderStyle = defaultBottomStyle, bottomColor: Color = defaultBottomColor): CellBorder = CellBorder(
    wrap(leftStyle, defaultLeftStyle), wrap(leftColor, defaultLeftColor),
    wrap(topStyle, defaultTopStyle), wrap(topColor, defaultTopColor),
    wrap(rightStyle, defaultRightStyle), wrap(rightColor, defaultRightColor),
    wrap(bottomStyle, defaultBottomStyle), wrap(bottomColor, defaultBottomColor)
  )
}

case class CellBorder(leftStyle: Option[BorderStyle], leftColor: Option[Color],
                      topStyle: Option[BorderStyle], topColor: Option[Color],
                      rightStyle: Option[BorderStyle], rightColor: Option[Color],
                      bottomStyle: Option[BorderStyle], bottomColor: Option[Color]) {

  def withLeftStyle(leftStyle: BorderStyle) =
    copy(leftStyle = Option(leftStyle))

  def withLeftColor(leftColor: Color) =
    copy(leftColor = Option(leftColor))

  def withTopStyle(topStyle: BorderStyle) =
    copy(topStyle = Option(topStyle))

  def withTopColor(topColor: Color) =
    copy(topColor = Option(topColor))

  def withRightStyle(rightStyle: BorderStyle) =
    copy(rightStyle = Option(rightStyle))

  def withRightColor(rightColor: Color) =
    copy(rightColor = Option(rightColor))

  def withBottomStyle(bottomStyle: BorderStyle) =
    copy(bottomStyle = Option(bottomStyle))

  def withBottomColor(bottomColor: Color) =
    copy(bottomColor = Option(bottomColor))

  def withStyle(style: BorderStyle) = {
    val styleOption = Option(style)
    copy(leftStyle = styleOption, topStyle = styleOption, rightStyle = styleOption, bottomStyle = styleOption)
  }

  def withColor(color: Color) = {
    val colorOption = Option(color)
    copy(leftColor = colorOption, topColor = colorOption, rightColor = colorOption, bottomColor = colorOption)
  }

  def applyTo(style: XSSFCellStyle) {
    leftStyle.foreach(style.setBorderLeft)
    leftColor.foreach(c => style.setLeftBorderColor(c.convert()))
    bottomStyle.foreach(style.setBorderBottom)
    bottomColor.foreach(c => style.setBottomBorderColor(c.convert()))
    rightStyle.foreach(style.setBorderRight)
    rightColor.foreach(c => style.setRightBorderColor(c.convert()))
    topStyle.foreach(style.setBorderTop)
    topColor.foreach(c => style.setTopBorderColor(c.convert()))
  }
}
