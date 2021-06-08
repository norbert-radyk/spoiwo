package spoiwo.model.enums

object CellBorderStyle {

  lazy val DashDot: CellBorderStyle = CellBorderStyle("DashDot")
  lazy val DashDotDot: CellBorderStyle = CellBorderStyle("DashDotDot")
  lazy val Dashed: CellBorderStyle = CellBorderStyle("Dashed")
  lazy val Dotted: CellBorderStyle = CellBorderStyle("Dotted")
  lazy val Double: CellBorderStyle = CellBorderStyle("Double")
  lazy val Hair: CellBorderStyle = CellBorderStyle("Hair")
  lazy val Medium: CellBorderStyle = CellBorderStyle("Medium")
  lazy val MediumDashDot: CellBorderStyle = CellBorderStyle("MediumDashDot")
  lazy val MediumDashed: CellBorderStyle = CellBorderStyle("MediumDashed")
  lazy val None: CellBorderStyle = CellBorderStyle("None")
  lazy val SlantedDashDot: CellBorderStyle = CellBorderStyle("SlantedDashDot")
  lazy val Thick: CellBorderStyle = CellBorderStyle("Thick")
  lazy val Thin: CellBorderStyle = CellBorderStyle("Thin")

}

case class CellBorderStyle private (value: String) {
  override def toString: String = value
}
