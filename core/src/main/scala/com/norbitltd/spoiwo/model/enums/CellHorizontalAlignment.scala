package com.norbitltd.spoiwo.model.enums

object CellHorizontalAlignment {

  lazy val Center = CellHorizontalAlignment("Center")
  lazy val CenterSelection = CellHorizontalAlignment("CenterSelection")
  lazy val Disturbed = CellHorizontalAlignment("Disturbed")
  lazy val Fill = CellHorizontalAlignment("Fill")
  lazy val General = CellHorizontalAlignment("General")
  lazy val Justify = CellHorizontalAlignment("Justify")
  lazy val Left = CellHorizontalAlignment("Left")
  lazy val Right = CellHorizontalAlignment("Right")

}

case class CellHorizontalAlignment private (value: String) {
  override def toString: String = value
}
