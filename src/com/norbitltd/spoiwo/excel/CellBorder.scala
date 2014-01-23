package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.xssf.usermodel.XSSFCellStyle

object CellBorder {

  lazy val Default = apply()

}

case class CellBorder(leftStyle: BorderStyle = BorderStyle.NONE, leftColor: Color = Color.BLACK,
                      topStyle: BorderStyle = BorderStyle.NONE, topColor: Color = Color.BLACK,
                      rightStyle: BorderStyle = BorderStyle.NONE, rightColor: Color = Color.BLACK,
                      bottomStyle: BorderStyle = BorderStyle.NONE, bottomColor: Color = Color.BLACK) {

  def applyTo(style: XSSFCellStyle) {
    style.setBorderLeft(leftStyle)
    style.setLeftBorderColor(leftColor.convert())
    style.setBorderBottom(bottomStyle)
    style.setBottomBorderColor(bottomColor.convert())
    style.setBorderRight(rightStyle)
    style.setRightBorderColor(rightColor.convert())
    style.setBorderTop(topStyle)
    style.setTopBorderColor(topColor.convert())
  }

  def withLeftStyle(leftStyle: BorderStyle) =
    copy(leftStyle = leftStyle)

  def withLeftColor(leftColor: Color) =
    copy(leftColor = leftColor)

  def withTopStyle(topStyle: BorderStyle) =
    copy(topStyle = topStyle)

  def withTopColor(topColor: Color) =
    copy(topColor = topColor)

  def withRightStyle(rightStyle: BorderStyle) =
    copy(rightStyle = rightStyle)

  def withRightColor(rightColor: Color) =
    copy(rightColor = rightColor)

  def withBottomStyle(bottomStyle: BorderStyle) =
    copy(bottomStyle = bottomStyle)

  def withBottomColor(bottomColor: Color) =
    copy(bottomColor = bottomColor)

}
