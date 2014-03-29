package com.norbitltd.spoiwo.model

object CellHorizontalAlignment {

  lazy val Center = CellHorizontalAlignment(1)
  lazy val CenterSelection = CellHorizontalAlignment(2)
  lazy val Disturbed = CellHorizontalAlignment(3)
  lazy val Fill = CellHorizontalAlignment(4)
  lazy val General = CellHorizontalAlignment(5)
  lazy val Justify = CellHorizontalAlignment(6)
  lazy val Left = CellHorizontalAlignment(7)
  lazy val Right = CellHorizontalAlignment(8)

}

case class CellHorizontalAlignment private(id : Int)

