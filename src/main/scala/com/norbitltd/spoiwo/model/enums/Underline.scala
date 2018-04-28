package com.norbitltd.spoiwo.model.enums

object Underline {

  private[model] lazy val Undefined = Underline("Undefined")
  lazy val Double = Underline("Double")
  lazy val DoubleAccounting = Underline("DoubleAccounting")
  lazy val None = Underline("None")
  lazy val Single = Underline("Single")
  lazy val SingleAccounting = Underline("SingleAccounting")

}

case class Underline private (value: String) {

  override def toString: String = value

}
