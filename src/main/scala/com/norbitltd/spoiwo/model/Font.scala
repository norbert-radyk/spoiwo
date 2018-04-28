package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.model.enums._

object Font {

  def apply(height: Height = null,
            bold: java.lang.Boolean = null,
            italic: java.lang.Boolean = null,
            charSet: Charset = null,
            color: Color = null,
            family: FontFamily = null,
            scheme: FontScheme = null,
            fontName: String = null,
            strikeout: java.lang.Boolean = null,
            typeOffset: TypeOffset = null,
            underline: Underline = null): Font =
    Font(
      height = Option(height),
      bold = Option(bold).map(_.booleanValue),
      italic = Option(italic).map(_.booleanValue),
      charSet = Option(charSet),
      color = Option(color),
      family = Option(family),
      scheme = Option(scheme),
      fontName = Option(fontName),
      strikeout = Option(strikeout).map(_.booleanValue),
      typeOffset = Option(typeOffset),
      underline = Option(underline)
    )
}

case class Font private[model] (height: Option[Height],
                                bold: Option[Boolean],
                                italic: Option[Boolean],
                                charSet: Option[Charset],
                                color: Option[Color],
                                family: Option[FontFamily],
                                scheme: Option[FontScheme],
                                fontName: Option[String],
                                strikeout: Option[Boolean],
                                typeOffset: Option[TypeOffset],
                                underline: Option[Underline]) {

  def withHeight(height: Height): Font =
    copy(height = Option(height))

  def withoutHeight: Font =
    copy(height = None)

  def withBold: Font =
    copy(bold = Some(true))

  def withoutBold: Font =
    copy(bold = Some(false))

  def withItalic: Font =
    copy(italic = Some(true))

  def withoutItalic: Font =
    copy(italic = Some(false))

  def withCharSet(charSet: Charset): Font =
    copy(charSet = Option(charSet))

  def withoutCharSet: Font =
    copy(charSet = None)

  def withColor(color: Color): Font =
    copy(color = Option(color))

  def withoutColor: Font =
    copy(color = None)

  def withFamily(family: FontFamily): Font =
    copy(family = Option(family))

  def withoutFamily: Font =
    copy(family = None)

  def withScheme(scheme: FontScheme): Font =
    copy(scheme = Option(scheme))

  def withoutScheme: Font =
    copy(scheme = None)

  def withFontName(fontName: String): Font =
    copy(fontName = Option(fontName))

  def withoutFontName: Font =
    copy(fontName = None)

  def withStrikeout: Font =
    copy(strikeout = Some(true))

  def withoutStrikeout: Font =
    copy(strikeout = Some(false))

  def withTypeOffset(typeOffset: TypeOffset): Font =
    copy(typeOffset = Option(typeOffset))

  def withoutTypeOffset: Font =
    copy(typeOffset = None)

  def withUnderline(underline: Underline): Font =
    copy(underline = Option(underline))

  def withoutUnderline: Font =
    copy(underline = None)

  def defaultWith(defaultFont: Font): Font = Font(
    height = dw(height, defaultFont.height),
    bold = dw(bold, defaultFont.bold),
    italic = dw(italic, defaultFont.italic),
    charSet = dw(charSet, defaultFont.charSet),
    color = dw(color, defaultFont.color),
    family = dw(family, defaultFont.family),
    scheme = dw(scheme, defaultFont.scheme),
    fontName = dw(fontName, defaultFont.fontName),
    strikeout = dw(strikeout, defaultFont.strikeout),
    typeOffset = dw(typeOffset, defaultFont.typeOffset),
    underline = dw(underline, defaultFont.underline)
  )

  private def dw[T](current: Option[T], default: Option[T]): Option[T] =
    if (current.isDefined) current else default

  override def toString: String =
    "Font[" + List(
      height.map("height=" + _),
      bold.map("bold=" + _),
      italic.map("italic" + _),
      charSet.map("charset=" + _),
      color.map("color=" + _),
      family.map("family=" + _),
      scheme.map("scheme=" + _),
      fontName.map("font name=" + _),
      strikeout.map("strikeout=" + _),
      typeOffset.map("type offset=" + _),
      underline.map("underline=" + _)
    ).flatten.mkString(", ") + "]"

}
