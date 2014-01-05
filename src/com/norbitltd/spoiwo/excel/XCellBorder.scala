package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.usermodel.BorderStyle

object XCellBorder {

  def apply(borderStyle : BorderStyle) : XCellBorder = XCellBorder(borderStyle, borderStyle, borderStyle, borderStyle)

}

case class XCellBorder(left : BorderStyle, top : BorderStyle, right : BorderStyle, bottom : BorderStyle) {

}
