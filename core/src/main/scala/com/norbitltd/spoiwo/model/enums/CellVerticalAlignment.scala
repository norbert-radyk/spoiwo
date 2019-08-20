package com.norbitltd.spoiwo.model.enums

object CellVerticalAlignment {

  lazy val Bottom = CellVerticalAlignment("Bottom")
  lazy val Center = CellVerticalAlignment("Center")
  lazy val Disturbed = CellVerticalAlignment("Disturbed")
  lazy val Justify = CellVerticalAlignment("Justify")
  lazy val Top = CellVerticalAlignment("Top")

}

case class CellVerticalAlignment private (value: String) {
  override def toString: String = value
}
