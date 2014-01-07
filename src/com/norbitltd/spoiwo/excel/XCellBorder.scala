package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.usermodel.{CellStyle, BorderStyle}
import org.apache.poi.xssf.usermodel.XSSFCellStyle

object XCellBorder {

  def apply(borderStyle: BorderStyle, borderColor: XColor = XColor.BLACK): XCellBorder =
    XCellBorder(borderStyle, borderColor, borderStyle, borderColor, borderStyle, borderColor, borderStyle, borderColor)
}

case class XCellBorder(leftStyle: BorderStyle, leftColor: XColor,
                       topStyle: BorderStyle, topColor: XColor,
                       rightStyle: BorderStyle, rightColor: XColor,
                       bottomStyle: BorderStyle, bottomColor: XColor) {

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
