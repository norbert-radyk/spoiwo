package spoiwo.model.enums

object CellVerticalAlignment {

  lazy val Bottom: CellVerticalAlignment = CellVerticalAlignment("Bottom")
  lazy val Center: CellVerticalAlignment = CellVerticalAlignment("Center")
  lazy val Disturbed: CellVerticalAlignment = CellVerticalAlignment("Disturbed")
  lazy val Justify: CellVerticalAlignment = CellVerticalAlignment("Justify")
  lazy val Top: CellVerticalAlignment = CellVerticalAlignment("Top")

}

case class CellVerticalAlignment private (value: String) {
  override def toString: String = value
}
