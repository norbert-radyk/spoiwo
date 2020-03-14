package com.norbitltd.spoiwo.model.enums

object Underline {

  private[model] lazy val Undefined = Underline("Undefined")
  lazy val Double: Underline = Underline("Double")
  lazy val DoubleAccounting: Underline = Underline("DoubleAccounting")
  lazy val None: Underline = Underline("None")
  lazy val Single: Underline = Underline("Single")
  lazy val SingleAccounting: Underline = Underline("SingleAccounting")

}

case class Underline private (value: String) {

  override def toString: String = value

}
