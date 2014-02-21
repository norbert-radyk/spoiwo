package com.norbitltd.spoiwo.model

object CellBorderStyle {

  lazy val DashDot = CellBorderStyle(1)
  lazy val DashDotDot = CellBorderStyle(2)
  lazy val Dashed = CellBorderStyle(3)
  lazy val Dotted = CellBorderStyle(4)
  lazy val Double = CellBorderStyle(5)
  lazy val Hair = CellBorderStyle(6)
  lazy val Medium = CellBorderStyle(7)
  lazy val MediumDashDot = CellBorderStyle(8)
  lazy val MediumDashDotDot = CellBorderStyle(9)
  lazy val MediumDashed = CellBorderStyle(10)
  lazy val None = CellBorderStyle(11)
  lazy val SlantedDashDot = CellBorderStyle(12)
  lazy val Thick = CellBorderStyle(13)
  lazy val Thin = CellBorderStyle(14)

}

case class CellBorderStyle private(id : Int)

