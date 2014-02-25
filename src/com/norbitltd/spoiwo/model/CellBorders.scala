package com.norbitltd.spoiwo.model

object CellBorders extends Factory {

  private lazy val defaultLeftStyle = CellBorderStyle.None
  private lazy val defaultLeftColor = Color.Undefined
  private lazy val defaultTopStyle = CellBorderStyle.None
  private lazy val defaultTopColor = Color.Undefined
  private lazy val defaultRightStyle = CellBorderStyle.None
  private lazy val defaultRightColor = Color.Undefined
  private lazy val defaultBottomStyle = CellBorderStyle.None
  private lazy val defaultBottomColor = Color.Undefined

  val Default = CellBorders()

  def apply(leftStyle: CellBorderStyle = defaultLeftStyle, leftColor: Color = defaultLeftColor,
            topStyle: CellBorderStyle = defaultTopStyle, topColor: Color = defaultTopColor,
            rightStyle: CellBorderStyle = defaultRightStyle, rightColor: Color = defaultRightColor,
            bottomStyle: CellBorderStyle = defaultBottomStyle, bottomColor: Color = defaultBottomColor): CellBorders = CellBorders(
    wrap(leftStyle, defaultLeftStyle), wrap(leftColor, defaultLeftColor),
    wrap(topStyle, defaultTopStyle), wrap(topColor, defaultTopColor),
    wrap(rightStyle, defaultRightStyle), wrap(rightColor, defaultRightColor),
    wrap(bottomStyle, defaultBottomStyle), wrap(bottomColor, defaultBottomColor)
  )
}

case class CellBorders(leftStyle: Option[CellBorderStyle], leftColor: Option[Color],
                      topStyle: Option[CellBorderStyle], topColor: Option[Color],
                      rightStyle: Option[CellBorderStyle], rightColor: Option[Color],
                      bottomStyle: Option[CellBorderStyle], bottomColor: Option[Color]) {

  def withLeftStyle(leftStyle: CellBorderStyle) =
    copy(leftStyle = Option(leftStyle))

  def withLeftColor(leftColor: Color) =
    copy(leftColor = Option(leftColor))

  def withTopStyle(topStyle: CellBorderStyle) =
    copy(topStyle = Option(topStyle))

  def withTopColor(topColor: Color) =
    copy(topColor = Option(topColor))

  def withRightStyle(rightStyle: CellBorderStyle) =
    copy(rightStyle = Option(rightStyle))

  def withRightColor(rightColor: Color) =
    copy(rightColor = Option(rightColor))

  def withBottomStyle(bottomStyle: CellBorderStyle) =
    copy(bottomStyle = Option(bottomStyle))

  def withBottomColor(bottomColor: Color) =
    copy(bottomColor = Option(bottomColor))

  def withStyle(style: CellBorderStyle) = {
    val styleOption = Option(style)
    copy(leftStyle = styleOption, topStyle = styleOption, rightStyle = styleOption, bottomStyle = styleOption)
  }

  def withColor(color: Color) = {
    val colorOption = Option(color)
    copy(leftColor = colorOption, topColor = colorOption, rightColor = colorOption, bottomColor = colorOption)
  }

}
