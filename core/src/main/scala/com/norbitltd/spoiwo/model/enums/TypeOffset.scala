package com.norbitltd.spoiwo.model.enums

object TypeOffset {

  lazy val None: TypeOffset = TypeOffset("None")
  lazy val Subscript: TypeOffset = TypeOffset("Subscript")
  lazy val Superscript: TypeOffset = TypeOffset("Superscript")
}
case class TypeOffset private (value: String) {

  override def toString: String = value

}
