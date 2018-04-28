package com.norbitltd.spoiwo.model.enums

object CellBorderStyle {

  lazy val DashDot = CellBorderStyle("DashDot")
  lazy val DashDotDot = CellBorderStyle("DashDotDot")
  lazy val Dashed = CellBorderStyle("Dashed")
  lazy val Dotted = CellBorderStyle("Dotted")
  lazy val Double = CellBorderStyle("Double")
  lazy val Hair = CellBorderStyle("Hair")
  lazy val Medium = CellBorderStyle("Medium")
  lazy val MediumDashDot = CellBorderStyle("MediumDashDot")
  lazy val MediumDashed = CellBorderStyle("MediumDashed")
  lazy val None = CellBorderStyle("None")
  lazy val SlantedDashDot = CellBorderStyle("SlantedDashDot")
  lazy val Thick = CellBorderStyle("Thick")
  lazy val Thin = CellBorderStyle("Thin")

}

case class CellBorderStyle private (value: String) {
  override def toString: String = value
}
