package com.norbitltd.spoiwo.model.enums

object CellReadingOrder {

  lazy val Context: CellReadingOrder = CellReadingOrder("Context")
  lazy val LeftToRight: CellReadingOrder = CellReadingOrder("LeftToRight")
  lazy val RightToLeft: CellReadingOrder = CellReadingOrder("RightToLeft")

}

case class CellReadingOrder private (value: String) {
  override def toString: String = value
}
