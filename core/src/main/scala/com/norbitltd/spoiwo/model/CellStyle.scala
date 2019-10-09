package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.model.enums.{CellReadingOrder, CellVerticalAlignment, CellHorizontalAlignment, CellFill}

object CellStyle {

  val Default = CellStyle()

  def apply(borders: CellBorders = null,
            dataFormat: CellDataFormat = null,
            font: Font = null,
            fillPattern: CellFill = null,
            fillForegroundColor: Color = null,
            fillBackgroundColor: Color = null,
            readingOrder: CellReadingOrder = null,
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
      readingOrder = Option(readingOrder),
      horizontalAlignment = Option(horizontalAlignment),
      verticalAlignment = Option(verticalAlignment),
      hidden = Option(hidden).map(_.booleanValue),
      indention = Option(indention).map(_.shortValue),
      locked = Option(locked).map(_.booleanValue),
      rotation = Option(rotation).map(_.shortValue),
      wrapText = Option(wrapText).map(_.booleanValue)
    )
}

case class CellStyle private (borders: Option[CellBorders],
                              dataFormat: Option[CellDataFormat],
                              font: Option[Font],
                              fillPattern: Option[CellFill],
                              fillForegroundColor: Option[Color],
                              fillBackgroundColor: Option[Color],
                              readingOrder: Option[CellReadingOrder],
                              horizontalAlignment: Option[CellHorizontalAlignment],
                              verticalAlignment: Option[CellVerticalAlignment],
                              hidden: Option[Boolean],
                              indention: Option[Short],
                              locked: Option[Boolean],
                              rotation: Option[Short],
                              wrapText: Option[Boolean]) {

  override def toString: String =
    "CellStyle(" + List(
      borders.map("borders=" + _),
      dataFormat.map("data format=" + _),
      font.map("font=" + _),
      fillPattern.map("fill pattern=" + _),
      fillForegroundColor.map("fill foreground color=" + _),
      fillBackgroundColor.map("fill background color=" + _),
      readingOrder.map("reading order=" + _),
      horizontalAlignment.map("horizontal alignment=" + _),
      verticalAlignment.map("vertical alignment=" + _),
      hidden.map("hidden=" + _),
      indention.map("indention=" + _),
      locked.map("locked=" + _),
      rotation.map("rotation=" + _),
      wrapText.map("wrap text=" + _)
    ).flatten.mkString(", ") + ")"

  def withBorders(borders: CellBorders): CellStyle =
    copy(borders = Option(borders))

  def withoutBorders: CellStyle =
    copy(borders = None)

  def withDataFormat(dataFormat: CellDataFormat): CellStyle =
    copy(dataFormat = Option(dataFormat))

  def withoutDataFormat: CellStyle =
    copy(dataFormat = None)

  def withFont(font: Font): CellStyle =
    copy(font = Option(font))

  def withoutFont: CellStyle =
    copy(font = None)

  def withFillPattern(fillPattern: CellFill): CellStyle =
    copy(fillPattern = Option(fillPattern))

  def withoutFillPattern: CellStyle =
    copy(fillPattern = None)

  def withFillForegroundColor(fillForegroundColor: Color): CellStyle =
    copy(fillForegroundColor = Option(fillForegroundColor))

  def withoutFillForegroundColor: CellStyle =
    copy(fillForegroundColor = None)

  def withFillBackgroundColor(fillBackgroundColor: Color): CellStyle =
    copy(fillBackgroundColor = Option(fillBackgroundColor))

  def withoutFillBackgroundColor: CellStyle =
    copy(fillBackgroundColor = None)

  def withReadingOrder(readingOrder: CellReadingOrder): CellStyle =
    copy(readingOrder = Option(readingOrder))

  def withoutReadingOrder: CellStyle =
    copy(readingOrder = None)

  def withHorizontalAlignment(horizontalAlignment: CellHorizontalAlignment): CellStyle =
    copy(horizontalAlignment = Option(horizontalAlignment))

  def withoutHorizontalAlignment: CellStyle =
    copy(horizontalAlignment = None)

  def withVerticalAlignment(verticalAlignment: CellVerticalAlignment): CellStyle =
    copy(verticalAlignment = Option(verticalAlignment))

  def withoutVerticalAlignment: CellStyle =
    copy(verticalAlignment = None)

  def withHidden: CellStyle =
    copy(hidden = Some(true))

  def withoutHidden: CellStyle =
    copy(hidden = Some(false))

  def withIndention(indention: Short): CellStyle =
    copy(indention = Option(indention))

  def withoutIndention: CellStyle =
    copy(indention = None)

  def withLocked: CellStyle =
    copy(locked = Some(true))

  def withoutLocked: CellStyle =
    copy(locked = Some(false))

  def withRotation(rotation: Short): CellStyle =
    copy(rotation = Option(rotation))

  def withoutRotation: CellStyle =
    copy(rotation = None)

  def withWrapText: CellStyle =
    copy(wrapText = Some(true))

  def withoutWrapText: CellStyle =
    copy(wrapText = Some(false))

  private def dw[T](current: Option[T], default: Option[T]): Option[T] =
    if (current.isDefined) current else default

  private def defaultFont(defaultCellStyle: CellStyle): Option[Font] =
    if (defaultCellStyle.font.isEmpty) {
      font
    } else if (font.isEmpty) {
      defaultCellStyle.font
    } else {
      Option(font.get.defaultWith(defaultCellStyle.font.get))
    }

  def defaultWith(defaultCellStyle: CellStyle): CellStyle = CellStyle(
    borders = dw(borders, defaultCellStyle.borders),
    dataFormat = dw(dataFormat, defaultCellStyle.dataFormat),
    font = defaultFont(defaultCellStyle),
    fillPattern = dw(fillPattern, defaultCellStyle.fillPattern),
    fillForegroundColor = dw(fillForegroundColor, defaultCellStyle.fillForegroundColor),
    fillBackgroundColor = dw(fillBackgroundColor, defaultCellStyle.fillBackgroundColor),
    readingOrder = dw(readingOrder, defaultCellStyle.readingOrder),
    horizontalAlignment = dw(horizontalAlignment, defaultCellStyle.horizontalAlignment),
    verticalAlignment = dw(verticalAlignment, defaultCellStyle.verticalAlignment),
    hidden = dw(hidden, defaultCellStyle.hidden),
    indention = dw(indention, defaultCellStyle.indention),
    locked = dw(locked, defaultCellStyle.locked),
    rotation = dw(rotation, defaultCellStyle.rotation),
    wrapText = dw(wrapText, defaultCellStyle.wrapText)
  )

}
