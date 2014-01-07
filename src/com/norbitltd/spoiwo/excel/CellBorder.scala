package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.xssf.usermodel.XSSFCellStyle

object CellBorder {

  lazy val Default = apply()

  def apply(borderStyle: BorderStyle, borderColor: Color = Color.BLACK): CellBorder =
    CellBorder(borderStyle, borderColor, borderStyle, borderColor, borderStyle, borderColor, borderStyle, borderColor)

}

case class CellBorder(leftStyle: BorderStyle = BorderStyle.NONE, leftColor: Color = Color.BLACK,
                       topStyle: BorderStyle = BorderStyle.NONE, topColor: Color = Color.BLACK,
                       rightStyle: BorderStyle = BorderStyle.NONE, rightColor: Color = Color.BLACK,
                       bottomStyle: BorderStyle = BorderStyle.NONE, bottomColor: Color = Color.BLACK) {

  def convert(style: XSSFCellStyle) {
    style.setBorderLeft(leftStyle)
    style.setLeftBorderColor(leftColor.convert())
    style.setBorderBottom(bottomStyle)
    style.setBottomBorderColor(bottomColor.convert())
    style.setBorderRight(rightStyle)
    style.setRightBorderColor(rightColor.convert())
    style.setBorderTop(topStyle)
    style.setTopBorderColor(topColor.convert())
  }

}
