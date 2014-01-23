package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.{XSSFFont, XSSFWorkbook}
import org.apache.poi.xssf.model.ThemesTable

object Font {

  private[excel] val cache = collection.mutable.Map[Font, XSSFFont]()

}

case class Font(
                 height: Double = XSSFFont.DEFAULT_FONT_SIZE,
                 bold: Boolean = false,
                 italic: Boolean = false,
                 charSet: FontCharset = FontCharset.DEFAULT,
                 color: Color = Color.BLACK,
                 family: FontFamily = FontFamily.MODERN,
                 scheme: FontScheme = FontScheme.NONE,
                 fontName: String = XSSFFont.DEFAULT_FONT_NAME,
                 strikeout: Boolean = false,
                 typeOffset: Short = 0,
                 underline: FontUnderline = FontUnderline.NONE) {

  def convert(workbook: XSSFWorkbook): XSSFFont =
    Font.cache.getOrElseUpdate(this, createFont(workbook))

  private def createFont(workbook: XSSFWorkbook): XSSFFont = {
    val font = workbook.createFont()
    font.setBold(bold)
    font.setCharSet(charSet)
    font.setColor(color.convert())
    font.setFamily(family)
    font.setFontHeight(height)
    font.setItalic(italic)
    font.setScheme(scheme)
    font.setFontName(fontName)
    font.setStrikeout(strikeout)
    font.setTypeOffset(typeOffset)
    font.setUnderline(underline)
    font
  }

  def withHeight(height: Double) =
    copy(height = height)

  def withBold(bold: Boolean) =
    copy(bold = bold)

  def withItalic(italic: Boolean) =
    copy(italic = italic)

  def withCharSet(charSet: FontCharset) =
    copy(charSet = charSet)

  def withColor(color: Color) =
    copy(color = color)

  def withFamily(family: FontFamily) =
    copy(family = family)

  def withScheme(scheme: FontScheme) =
    copy(scheme = scheme)

  def withFontName(fontName: String) =
    copy(fontName = fontName)

  def withStrikeout(strikeout: Boolean) =
    copy(strikeout = strikeout)

  def withTypeOffset(typeOffset: Short) =
    copy(typeOffset = typeOffset)

  def withUnderline(underline: FontUnderline) =
    copy(underline = underline)

}
