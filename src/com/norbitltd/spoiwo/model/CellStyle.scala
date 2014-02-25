package com.norbitltd.spoiwo.model

object CellStyle extends Factory {

  private lazy val defaultDataFormat = CellDataFormat.Undefined
  private lazy val defaultFillPattern = CellFill.None
  private lazy val defaultFillForegroundColor = Color.Undefined
  private lazy val defaultFillBackgroundColor = Color.Undefined
  private lazy val defaultHorizontalAlignment = CellHorizontalAlignment.Undefined
  private lazy val defaultVerticalAlignment = CellVerticalAlignment.Undefined
  private lazy val defaultHidden = false
  private lazy val defaultIndention = -1.toShort
  private lazy val defaultLocked = false
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
      wrap(fillPattern, fillPattern),
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

  def withBorders(borders: CellBorders) =
    copy(borders = Option(borders))

  def withDataFormat(dataFormat: CellDataFormat) =
    copy(dataFormat = Option(dataFormat))

  def withFont(font: Font) =
    copy(font = Option(font))

  def withFillPattern(fillPattern: CellFill) =
    copy(fillPattern = Option(fillPattern))

  def withFillForegroundColor(fillForegroundColor: Color) =
    copy(fillForegroundColor = Option(fillForegroundColor))

  def withFillBackgroundColor(fillBackgroundColor: Color) =
    copy(fillBackgroundColor = Option(fillBackgroundColor))

  def withHorizontalAlignment(horizontalAlignment: CellHorizontalAlignment) =
    copy(horizontalAlignment = Option(horizontalAlignment))

  def withVerticalAlignment(verticalAlignment: CellVerticalAlignment) =
    copy(verticalAlignment = Option(verticalAlignment))

  def withHidden(hidden: Boolean) =
    copy(hidden = Option(hidden))

  def withIndention(indention: Short) =
    copy(indention = Option(indention))

  def withLocked(locked: Boolean) =
    copy(locked = Option(locked))

  def withRotation(rotation: Short) =
    copy(rotation = Option(rotation))

  def withWrapText(wrapText: Boolean) =
    copy(wrapText = Option(wrapText))

}
