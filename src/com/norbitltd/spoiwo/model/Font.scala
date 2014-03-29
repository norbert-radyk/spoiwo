package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.model.enums._
import com.norbitltd.spoiwo.utils.BooleanExt

object Font extends Factory {

  val Default = Font()

  def apply(height: Height = null, bold: BooleanExt = BooleanExt.Undefined, italic: BooleanExt = BooleanExt.Undefined,
            charSet: Charset = null, color: Color = null, family: FontFamily = null, scheme: FontScheme = null,
            fontName: String = null, strikeout: BooleanExt = BooleanExt.Undefined, typeOffset: TypeOffset = null,
            underline: Underline = null): Font =
    Font(
      height = Option(height),
      bold = bold.toOption,
      italic = italic.toOption,
      charSet = Option(charSet),
      color = Option(color),
      family = Option(family),
      scheme = Option(scheme),
      fontName = Option(fontName),
      strikeout = strikeout.toOption,
      typeOffset = Option(typeOffset),
      underline = Option(underline)
    )
}

case class Font private[model](
                                height: Option[Height],
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

  def withHeight(height: Height) =
    copy(height = Option(height))

  def withoutHeight =
    copy(height = None)

  def withBold =
    copy(bold = Some(true))

  def withoutBold =
    copy(bold = Some(false))

  def withItalic =
    copy(italic = Some(true))

  def withoutItalic =
    copy(italic = Some(false))

  def withCharSet(charSet: Charset) =
    copy(charSet = Option(charSet))

  def withoutCharSet =
    copy(charSet = None)

  def withColor(color: Color) =
    copy(color = Option(color))

  def withoutColor =
    copy(color = None)

  def withFamily(family: FontFamily) =
    copy(family = Option(family))

  def withoutFamily =
    copy(family = None)

  def withScheme(scheme: FontScheme) =
    copy(scheme = Option(scheme))

  def withoutScheme =
    copy(scheme = None)

  def withFontName(fontName: String) =
    copy(fontName = Option(fontName))

  def withoutFontName =
    copy(fontName = None)

  def withStrikeout =
    copy(strikeout = Some(true))

  def withoutStrikeout =
    copy(strikeout = Some(false))

  def withTypeOffset(typeOffset: TypeOffset) =
    copy(typeOffset = Option(typeOffset))

  def withoutTypeOffset =
    copy(typeOffset = None)

  def withUnderline(underline: Underline) =
    copy(underline = Option(underline))

  def withoutUnderline =
    copy(underline = None)

  def withDefault(default : Font) : Font = Font(
    height = setWithDefault(height, default.height),
    bold = setWithDefault(bold, default.bold),
    italic = setWithDefault(italic, default.italic),
    charSet = setWithDefault(charSet, default.charSet),
    color = setWithDefault(color, default.color),
    family = setWithDefault(family, default.family),
    scheme = setWithDefault(scheme, default.scheme),
    fontName = setWithDefault(fontName, default.fontName),
    strikeout = setWithDefault(strikeout, default.strikeout),
    typeOffset = setWithDefault(typeOffset, default.typeOffset),
    underline = setWithDefault(underline, default.underline)
  )

  override def toString = "Font[" + List(
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

  private def setWithDefault[T](value: Option[T], defaultValue: Option[T]): Option[T] =
    if (value.isDefined) value else defaultValue
}
