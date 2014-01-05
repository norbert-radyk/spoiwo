package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

object XFont {

  private[XFont] val cache = collection.mutable.Map[XFont, Font]()

}

case class XFont(
                  heightInPoints: Short = 9,
                  bold: Boolean = false,
                  italic: Boolean = false,
                  fontCharset: FontCharset = FontCharset.DEFAULT,
                  fontColor: XColor = XColor.BLACK,
                  fontFamily: FontFamily = FontFamily.MODERN,
                  fontScheme: FontScheme = FontScheme.NONE,
                  strikeout: Boolean = false,
                  underline: FontUnderline = FontUnderline.NONE) {

  def convert(workbook: XSSFWorkbook): Font =
    XFont.cache.getOrElseUpdate(this, createFont(workbook))

  private def createFont(workbook: XSSFWorkbook): Font = {
    val font = workbook.createFont()
    font.setBold(bold)
    font.setCharSet(fontCharset)
    font.setColor(fontColor.convert())
    font.setFamily(fontFamily)
    font.setFontHeightInPoints(heightInPoints)
    font.setItalic(italic)
    font.setScheme(fontScheme)
    font.setStrikeout(strikeout)
    font.setUnderline(underline)
    font
  }

}
