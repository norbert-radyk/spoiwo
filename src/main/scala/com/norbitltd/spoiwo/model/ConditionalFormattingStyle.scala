package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.model.enums.{CellBorderStyle, CellFill, Underline, TypeOffset}

object ConditionalFormattingStyle {

  def apply(fontTypeOffset: TypeOffset = null,
            fontColorIndex: java.lang.Short = null,
            fontHeight: Height = null,
            fontBold: java.lang.Boolean = null,
            fontItalic: java.lang.Boolean = null,
            fontUnderline: Underline = null,
            fillPattern: CellFill = null,
            fillForegroundColorIndex: java.lang.Short = null,
            fillBackgroundColorIndex: java.lang.Short = null,
            borderLeftStyle: CellBorderStyle = null,
            borderLeftColorIndex: java.lang.Short = null,
            borderTopStyle: CellBorderStyle = null,
            borderTopColorIndex: java.lang.Short = null,
            borderRightStyle: CellBorderStyle = null,
            borderRightColorIndex: java.lang.Short = null,
            borderBottomStyle: CellBorderStyle = null,
            borderBottomColorIndex: java.lang.Short = null,
            borderDiagonalStyle: CellBorderStyle = null,
            borderDiagonalColorIndex: java.lang.Short = null): ConditionalFormattingStyle = new ConditionalFormattingStyle(
    Option(fontTypeOffset),
    Option(fontColorIndex),
    Option(fontHeight),
    Option(fontBold).map(_.booleanValue),
    Option(fontItalic).map(_.booleanValue),
    Option(fontUnderline),
    Option(fillPattern),
    Option(fillForegroundColorIndex).map(_.shortValue),
    Option(fillBackgroundColorIndex).map(_.shortValue),
    Option(borderLeftStyle),
    Option(borderLeftColorIndex).map(_.shortValue),
    Option(borderTopStyle),
    Option(borderTopColorIndex).map(_.shortValue),
    Option(borderRightStyle),
    Option(borderRightColorIndex).map(_.shortValue),
    Option(borderBottomStyle),
    Option(borderBottomColorIndex).map(_.shortValue),
    Option(borderDiagonalStyle),
    Option(borderDiagonalColorIndex).map(_.shortValue)
  )

}

case class ConditionalFormattingStyle private[model](
                                                      fontTypeOffset: Option[TypeOffset],
                                                      fontColorIndex: Option[Short],
                                                      fontHeight: Option[Height],
                                                      fontBold: Option[Boolean],
                                                      fontItalic: Option[Boolean],
                                                      fontUnderline: Option[Underline],
                                                      fillPattern: Option[CellFill],
                                                      fillForegroundColorIndex: Option[Short],
                                                      fillBackgroundColorIndex: Option[Short],
                                                      borderLeftStyle: Option[CellBorderStyle],
                                                      borderLeftColorIndex: Option[Short],
                                                      borderTopStyle: Option[CellBorderStyle],
                                                      borderTopColorIndex: Option[Short],
                                                      borderRightStyle: Option[CellBorderStyle],
                                                      borderRightColorIndex: Option[Short],
                                                      borderBottomStyle: Option[CellBorderStyle],
                                                      borderBottomColorIndex: Option[Short],
                                                      borderDiagonalStyle: Option[CellBorderStyle],
                                                      borderDiagonalColorIndex: Option[Short]) {

  def withFontTypeOffset(fontTypeOffset: TypeOffset) =
    copy(fontTypeOffset = Option(fontTypeOffset))

  def withoutFontTypeOffset =
    copy(fontTypeOffset = None)

  def withFontColorIndex(fontColorIndex : Short) =
    copy(fontColorIndex = Option(fontColorIndex))

  def withoutFontColorIndex =
    copy(fontColorIndex = None)

  def withFontHeight(fontHeight : Height) =
    copy(fontHeight = Option(fontHeight))

  def withoutFontHeight =
    copy(fontHeight = None)

  def withFontBold =
    copy(fontBold = Option(true))

  def withoutFontBold =
    copy(fontBold = Option(false))

  def withFontItalic =
    copy(fontItalic = Option(true))

  def withoutFontItalic =
    copy(fontItalic = Option(false))

  def withFontUnderline(fontUnderline : Underline) =
    copy(fontUnderline = Option(fontUnderline))

  def withoutFontUnderline =
    copy(fontUnderline = None)
  
  def withFillPattern(fillPattern : CellFill) = 
    copy(fillPattern = Option(fillPattern))

  def withoutFillPattern =
    copy(fillPattern = None)
  
  def withFillForegroundColorIndex(fillForegroundColorIndex : Short) = 
    copy(fillForegroundColorIndex = Option(fillForegroundColorIndex))

  def withoutFillForegroundColorIndex =
    copy(fillForegroundColorIndex = None)

  def withFillBackgroundColorIndex(fillBackgroundColorIndex : Short) =
    copy(fillBackgroundColorIndex = Option(fillBackgroundColorIndex))

  def withoutFillBackgroundColorIndex =
    copy(fillBackgroundColorIndex = None)

  def withBorderLeftStyle(borderLeftStyle : CellBorderStyle) =
    copy(borderLeftStyle = Option(borderLeftStyle))

  def withoutBorderLeftStyle =
    copy(borderLeftStyle = None)

  def withBorderLeftColorIndex(borderLeftColorIndex : Short) =
    copy(borderLeftColorIndex = Option(borderLeftColorIndex))

  def withoutBorderLeftColorIndex =
    copy(borderLeftColorIndex = None)

  def withBorderTopStyle(borderTopStyle : CellBorderStyle) =
    copy(borderTopStyle = Option(borderTopStyle))

  def withoutBorderTopStyle =
    copy(borderTopStyle = None)

  def withBorderTopColorIndex(borderTopColorIndex : Short) =
    copy(borderTopColorIndex = Option(borderTopColorIndex))

  def withoutBorderTopColorIndex =
    copy(borderTopColorIndex = None)

  def withBorderRightStyle(borderRightStyle : CellBorderStyle) =
    copy(borderRightStyle = Option(borderRightStyle))

  def withoutBorderRightStyle =
    copy(borderRightStyle = None)

  def withBorderRightColorIndex(borderRightColorIndex : Short) =
    copy(borderRightColorIndex = Option(borderRightColorIndex))

  def withoutBorderRightColorIndex =
    copy(borderRightColorIndex = None)

  def withBorderBottomStyle(borderBottomStyle : CellBorderStyle) =
    copy(borderBottomStyle = Option(borderBottomStyle))

  def withoutBorderBottomStyle =
    copy(borderBottomStyle = None)

  def withBorderBottomColorIndex(borderBottomColorIndex : Short) =
    copy(borderBottomColorIndex = Option(borderBottomColorIndex))

  def withoutBorderBottomColorIndex =
    copy(borderBottomColorIndex = None)

  def withBorderDiagonalStyle(borderDiagonalStyle : CellBorderStyle) =
    copy(borderDiagonalStyle = Option(borderDiagonalStyle))

  def withoutBorderDiagonalStyle =
    copy(borderDiagonalStyle = None)

  def withBorderDiagonalColorIndex(borderDiagonalColorIndex : Short) =
    copy(borderDiagonalColorIndex = Option(borderDiagonalColorIndex))

  def withoutBorderDiagonalColorIndex =
    copy(borderDiagonalColorIndex = None)

}
