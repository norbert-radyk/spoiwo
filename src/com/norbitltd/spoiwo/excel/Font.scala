package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.{XSSFFont, XSSFWorkbook}

object Font {

  private[Font] val cache = collection.mutable.Map[Font, XSSFFont]()

}

case class Font(
                  heightInPoints: Short = 9,
                  bold: Boolean = false,
                  italic: Boolean = false,
                  fontCharset: FontCharset = FontCharset.DEFAULT,
                  fontColor: Color = Color.BLACK,
                  fontFamily: FontFamily = FontFamily.MODERN,
                  fontScheme: FontScheme = FontScheme.NONE,
                  strikeout: Boolean = false,
                  underline: FontUnderline = FontUnderline.NONE) {

  def convert(workbook: XSSFWorkbook): XSSFFont =
    Font.cache.getOrElseUpdate(this, createFont(workbook))

  private def createFont(workbook: XSSFWorkbook): XSSFFont = {
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
