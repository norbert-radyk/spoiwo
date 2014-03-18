package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.model.enums.CellFill

object CellStyle extends Factory {

  private lazy val defaultDataFormat = CellDataFormat.Undefined
  private lazy val defaultFillPattern = CellFill.Undefined
  private lazy val defaultFillForegroundColor = Color.Undefined
  private lazy val defaultFillBackgroundColor = Color.Undefined
  private lazy val defaultHorizontalAlignment = CellHorizontalAlignment.Undefined
  private lazy val defaultVerticalAlignment = CellVerticalAlignment.Undefined
  private lazy val defaultHidden = false
  private lazy val defaultIndention = -1.toShort
  private lazy val defaultLocked = true
  private lazy val defaultRotation = -1.toShort
  private lazy val defaultWrapText = false

  private lazy val defaultBorders = CellBorders.Default
  private lazy val defaultFont = Font.Default

  val Default = CellStyle()

  def apply(borders: CellBorders = defaultBorders,
            dataFormat: CellDataFormat = defaultDataFormat,
            font: Font = defaultFont,
            fillPattern: CellFill = defaultFillPattern,
            fillForegroundColor: Color = defaultFillForegroundColor,
            fillBackgroundColor: Color = defaultFillBackgroundColor,
            horizontalAlignment: CellHorizontalAlignment = defaultHorizontalAlignment,
            verticalAlignment: CellVerticalAlignment = defaultVerticalAlignment,
            hidden: Boolean = defaultHidden,
            indention: Short = defaultIndention,
            locked: Boolean = defaultLocked,
            rotation: Short = defaultRotation,
            wrapText: Boolean = defaultWrapText): CellStyle =
    CellStyle(
      wrap(borders, defaultBorders),
      wrap(dataFormat, defaultDataFormat),
      wrap(font, defaultFont),
      wrap(fillPattern, defaultFillPattern),
      wrap(fillForegroundColor, defaultFillForegroundColor),
      wrap(fillBackgroundColor, defaultFillBackgroundColor),
      wrap(horizontalAlignment, defaultHorizontalAlignment),
      wrap(verticalAlignment, defaultVerticalAlignment),
      wrap(hidden, defaultHidden),
      wrap(indention, defaultIndention),
      wrap(locked, defaultLocked),
      wrap(rotation, defaultRotation),
      wrap(wrapText, defaultWrapText)
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
