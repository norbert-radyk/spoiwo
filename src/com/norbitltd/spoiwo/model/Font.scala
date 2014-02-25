package com.norbitltd.spoiwo.model

import org.apache.poi.ss.usermodel.{FontCharset, FontFamily, FontUnderline, FontScheme}
import org.apache.poi.xssf.usermodel.{XSSFFont, XSSFWorkbook}
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._

object Font extends Factory {

  private lazy val defaultHeight = -1.toShort
  private lazy val defaultHeightInPoints = -1.toShort
  private lazy val defaultBold = false
  private lazy val defaultItalic = defaultPOIFont.getItalic
  private lazy val defaultCharSet = FontCharset.valueOf(defaultPOIFont.getCharSet)
  private lazy val defaultColor = Color.Undefined
  private lazy val defaultFamily = FontFamily.valueOf(defaultPOIFont.getFamily)
  private lazy val defaultScheme = defaultPOIFont.getScheme
  private lazy val defaultFontName = defaultPOIFont.getFontName
  private lazy val defaultStrikeout = defaultPOIFont.getStrikeout
  private lazy val defaultTypeOffset = defaultPOIFont.getTypeOffset
  private lazy val defaultUnderline = FontUnderline.valueOf(defaultPOIFont.getUnderline)

  val Default = Font()

  def apply(height: Short = defaultHeight,
            heightInPoints: Short = defaultHeightInPoints,
            bold: Boolean = defaultBold,
            italic: Boolean = defaultItalic,
            charSet: FontCharset = defaultCharSet,
            color: Color = defaultColor,
            family: FontFamily = defaultFamily,
            scheme: FontScheme = defaultScheme,
            fontName: String = defaultFontName,
            strikeout: Boolean = defaultStrikeout,
            typeOffset: Short = defaultTypeOffset,
            underline: FontUnderline = defaultUnderline) : Font = Font(
    height = wrap(height, defaultHeight),
      heightInPoints = wrap(heightInPoints, defaultHeightInPoints),
    bold = wrap(bold, defaultBold),
    italic = wrap(italic, defaultItalic),
    charSet = wrap(charSet, defaultCharSet),
    color = wrap(color, defaultColor),
    family = wrap(family, defaultFamily),
    scheme = wrap(scheme, defaultScheme),
    fontName = wrap(fontName, defaultFontName),
    strikeout = wrap(strikeout, defaultStrikeout),
    typeOffset = wrap(typeOffset, defaultTypeOffset),
    underline = wrap(underline, defaultUnderline)
  )

  private[model] val cache = collection.mutable.Map[XSSFWorkbook,
    collection.mutable.Map[Font, XSSFFont]]()

}

case class Font private[model](
                 height: Option[Short],
        heightInPoints: Option[Short],
                 bold: Option[Boolean],
                 italic: Option[Boolean],
                 charSet: Option[FontCharset],
                 color: Option[Color],
                 family: Option[FontFamily],
                 scheme: Option[FontScheme],
                 fontName: Option[String],
                 strikeout: Option[Boolean],
                 typeOffset: Option[Short],
                 underline: Option[FontUnderline]) {

  def withHeight(height: Short) =
    copy(height = Option(height))

  def withHeightInPoints(heightInPoints: Short) =
    copy(heightInPoints = Option(heightInPoints))

  def withBold(bold: Boolean) =
    copy(bold = Option(bold))

  def withItalic(italic: Boolean) =
    copy(italic = Option(italic))

  def withCharSet(charSet: FontCharset) =
    copy(charSet = Option(charSet))

  def withColor(color: Color) =
    copy(color = Option(color))

  def withFamily(family: FontFamily) =
    copy(family = Option(family))

  def withScheme(scheme: FontScheme) =
    copy(scheme = Option(scheme))

  def withFontName(fontName: String) =
    copy(fontName = Option(fontName))

  def withStrikeout(strikeout: Boolean) =
    copy(strikeout = Option(strikeout))

  def withTypeOffset(typeOffset: Short) =
    copy(typeOffset = Option(typeOffset))

  def withUnderline(underline: FontUnderline) =
    copy(underline = Option(underline))

  def convert(workbook: XSSFWorkbook): XSSFFont = {
    val workbookCache = Font.cache.getOrElseUpdate(workbook, collection.mutable.Map[Font, XSSFFont]())
    workbookCache.getOrElseUpdate(this, createFont(workbook))
  }


  private def createFont(workbook: XSSFWorkbook): XSSFFont = {
    val font = workbook.createFont()
    bold.foreach(font.setBold)
    charSet.foreach(font.setCharSet)
    color.foreach(c => font.setColor(c.convertAsXlsx()))
    family.foreach(font.setFamily)
    height.foreach(font.setFontHeight)
    heightInPoints.foreach(font.setFontHeightInPoints)
    italic.foreach(font.setItalic)
    scheme.foreach(font.setScheme)
    fontName.foreach(font.setFontName)
    strikeout.foreach(font.setStrikeout)
    typeOffset.foreach(font.setTypeOffset)
    underline.foreach(font.setUnderline)
    font
  }


}
