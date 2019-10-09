package com.norbitltd.spoiwo.model.enums

object CellReadingOrder {

  lazy val Context = CellReadingOrder("Context")
  lazy val LeftToRight = CellReadingOrder("LeftToRight")
  lazy val RightToLeft = CellReadingOrder("RightToLeft")

}

case class CellReadingOrder private(value: String) {
  override def toString: String = value
}
