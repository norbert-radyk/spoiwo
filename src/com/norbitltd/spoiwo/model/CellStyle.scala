package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.model.enums.{CellVerticalAlignment, CellHorizontalAlignment, CellFill}

object CellStyle {

  val Default = CellStyle()

  def apply(borders: CellBorders = null,
            dataFormat: CellDataFormat = null,
            font: Font = null,
            fillPattern: CellFill = null,
            fillForegroundColor: Color = null,
            fillBackgroundColor: Color = null,
            horizontalAlignment: CellHorizontalAlignment = null,
            verticalAlignment: CellVerticalAlignment = null,
            hidden: java.lang.Boolean = null,
            indention: java.lang.Integer = null,
            locked: java.lang.Boolean = null,
            rotation: java.lang.Integer = null,
            wrapText: java.lang.Boolean = null): CellStyle =
    CellStyle(
      borders = Option(borders),
      dataFormat = Option(dataFormat),
      font = Option(font),
      fillPattern = Option(fillPattern),
      fillForegroundColor = Option(fillForegroundColor),
      fillBackgroundColor = Option(fillBackgroundColor),
      horizontalAlignment = Option(horizontalAlignment),
      verticalAlignment = Option(verticalAlignment),
      hidden = Option(hidden).map(_.booleanValue),
      indention = Option(indention).map(_.shortValue),
      locked = Option(locked).map(_.booleanValue),
      rotation = Option(rotation).map(_.shortValue),
      wrapText = Option(wrapText).map(_.booleanValue)
    )
}

case class CellStyle private(
                              borders: Option[CellBorders],
                              dataFormat: Option[CellDataFormat],
                              font: Option[Font],
                              fillPattern: Option[CellFill],
                              fillForegroundColor: Option[Color],
                              fillBackgroundColor: Option[Color],
                              horizontalAlignment: Option[CellHorizontalAlignment],
                              verticalAlignment: Option[CellVerticalAlignment],
                              hidden: Option[Boolean],
                              indention: Option[Short],
                              locked: Option[Boolean],
                              rotation: Option[Short],
                              wrapText: Option[Boolean]) {

  override def toString = "CellStyle(" + List(
    borders.map("borders=" + _),
    dataFormat.map("data format=" + _),
    font.map("font=" + _),
    fillPattern.map("fill pattern=" + _),
    fillForegroundColor.map("fill foreground color=" + _),
    fillBackgroundColor.map("fill background color=" + _),
    horizontalAlignment.map("horizontal alignment=" + _),
    verticalAlignment.map("vertical alignment=" + _),
    hidden.map("hidden=" + _),
    indention.map("indention=" + _),
    locked.map("locked=" + _),
    rotation.map("rotation=" + _),
    wrapText.map("wrap text=" + _)
  ).flatten.mkString(", ") + ")"

  def withBorders(borders: CellBorders) =
    copy(borders = Option(borders))

  def withoutBorders =
    copy(borders = None)

  def withDataFormat(dataFormat: CellDataFormat) =
    copy(dataFormat = Option(dataFormat))

  def withoutDataFormat =
    copy(dataFormat = None)

  def withFont(font: Font) =
    copy(font = Option(font))

  def withoutFont =
    copy(font = None)

  def withFillPattern(fillPattern: CellFill) =
    copy(fillPattern = Option(fillPattern))

  def withoutFillPattern =
    copy(fillPattern = None)

  def withFillForegroundColor(fillForegroundColor: Color) =
    copy(fillForegroundColor = Option(fillForegroundColor))

  def withoutFillForegroundColor =
    copy(fillForegroundColor = None)

  def withFillBackgroundColor(fillBackgroundColor: Color) =
    copy(fillBackgroundColor = Option(fillBackgroundColor))

  def withoutFillBackgroundColor =
    copy(fillBackgroundColor = None)

  def withHorizontalAlignment(horizontalAlignment: CellHorizontalAlignment) =
    copy(horizontalAlignment = Option(horizontalAlignment))

  def withoutHorizontalAlignment =
    copy(horizontalAlignment = None)

  def withVerticalAlignment(verticalAlignment: CellVerticalAlignment) =
    copy(verticalAlignment = Option(verticalAlignment))

  def withoutVerticalAlignment =
    copy(verticalAlignment = None)

  def withHidden =
    copy(hidden = Some(true))

  def withoutHidden =
    copy(hidden = Some(false))

  def withIndention(indention: Short) =
    copy(indention = Option(indention))

  def withoutIndention =
    copy(indention = None)

  def withLocked =
    copy(locked = Some(true))

  def withoutLocked =
    copy(locked = Some(false))

  def withRotation(rotation: Short) =
    copy(rotation = Option(rotation))

  def withoutRotation =
    copy(rotation = None)

  def withWrapText =
    copy(wrapText = Some(true))

  def withoutWrapText =
    copy(wrapText = Some(false))


}
