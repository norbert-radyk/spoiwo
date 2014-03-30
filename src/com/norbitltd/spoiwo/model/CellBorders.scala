package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.model.enums.CellBorderStyle

object CellBorders {

  def apply(leftStyle: CellBorderStyle = null, leftColor: Color = null,
            topStyle: CellBorderStyle = null, topColor: Color = null,
            rightStyle: CellBorderStyle = null, rightColor: Color = null,
            bottomStyle: CellBorderStyle = null, bottomColor: Color = null): CellBorders =
    CellBorders(
      Option(leftStyle), Option(leftColor),
      Option(topStyle), Option(topColor),
      Option(rightStyle), Option(rightColor),
      Option(bottomStyle), Option(bottomColor)
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
