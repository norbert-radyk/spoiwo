package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.model.enums.CellBorderStyle

object CellBorders {

  def apply(leftStyle: CellBorderStyle = null,
            leftColor: Color = null,
            topStyle: CellBorderStyle = null,
            topColor: Color = null,
            rightStyle: CellBorderStyle = null,
            rightColor: Color = null,
            bottomStyle: CellBorderStyle = null,
            bottomColor: Color = null): CellBorders =
    CellBorders(
      Option(leftStyle),
      Option(leftColor),
      Option(topStyle),
      Option(topColor),
      Option(rightStyle),
      Option(rightColor),
      Option(bottomStyle),
      Option(bottomColor)
    )

}

case class CellBorders(leftStyle: Option[CellBorderStyle],
                       leftColor: Option[Color],
                       topStyle: Option[CellBorderStyle],
                       topColor: Option[Color],
                       rightStyle: Option[CellBorderStyle],
                       rightColor: Option[Color],
                       bottomStyle: Option[CellBorderStyle],
                       bottomColor: Option[Color]) {

  def withLeftStyle(leftStyle: CellBorderStyle): CellBorders =
    copy(leftStyle = Option(leftStyle))

  def withLeftColor(leftColor: Color): CellBorders =
    copy(leftColor = Option(leftColor))

  def withTopStyle(topStyle: CellBorderStyle): CellBorders =
    copy(topStyle = Option(topStyle))

  def withTopColor(topColor: Color): CellBorders =
    copy(topColor = Option(topColor))

  def withRightStyle(rightStyle: CellBorderStyle): CellBorders =
    copy(rightStyle = Option(rightStyle))

  def withRightColor(rightColor: Color): CellBorders =
    copy(rightColor = Option(rightColor))

  def withBottomStyle(bottomStyle: CellBorderStyle): CellBorders =
    copy(bottomStyle = Option(bottomStyle))

  def withBottomColor(bottomColor: Color): CellBorders =
    copy(bottomColor = Option(bottomColor))

  def withStyle(style: CellBorderStyle): CellBorders = {
    val styleOption = Option(style)
    copy(leftStyle = styleOption, topStyle = styleOption, rightStyle = styleOption, bottomStyle = styleOption)
  }

  def withColor(color: Color): CellBorders = {
    val colorOption = Option(color)
    copy(leftColor = colorOption, topColor = colorOption, rightColor = colorOption, bottomColor = colorOption)
  }

  override def toString: String = "Cell Borders[" + List(
    leftStyle.map("left style" + _),
    leftColor.map("left color" + _),
    topStyle.map("top style" + _),
    topColor.map("top color" + _),
    rightStyle.map("right style" + _),
    rightColor.map("right color" + _),
    bottomStyle.map("bottom style" + _),
    bottomColor.map("bottom color" + _)
  )

}
