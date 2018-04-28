package com.norbitltd.spoiwo.model.enums

object TypeOffset {

  lazy val None = TypeOffset("None")
  lazy val Subscript = TypeOffset("Subscript")
  lazy val Superscript = TypeOffset("Superscript")
}
case class TypeOffset private (value: String) {

  override def toString: String = value

}
