package spoiwo.model.enums

object CellHorizontalAlignment {

  lazy val Center: CellHorizontalAlignment = CellHorizontalAlignment("Center")
  lazy val CenterSelection: CellHorizontalAlignment = CellHorizontalAlignment("CenterSelection")
  lazy val Disturbed: CellHorizontalAlignment = CellHorizontalAlignment("Disturbed")
  lazy val Fill: CellHorizontalAlignment = CellHorizontalAlignment("Fill")
  lazy val General: CellHorizontalAlignment = CellHorizontalAlignment("General")
  lazy val Justify: CellHorizontalAlignment = CellHorizontalAlignment("Justify")
  lazy val Left: CellHorizontalAlignment = CellHorizontalAlignment("Left")
  lazy val Right: CellHorizontalAlignment = CellHorizontalAlignment("Right")

}

case class CellHorizontalAlignment private (value: String) {
  override def toString: String = value
}
