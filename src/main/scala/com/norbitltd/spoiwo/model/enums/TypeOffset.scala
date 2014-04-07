package com.norbitltd.spoiwo.model.enums

object TypeOffset {

  private[model] lazy val Undefined = TypeOffset("Undefined")
  lazy val None = TypeOffset("None")
  lazy val Subscript = TypeOffset("Subscript")
  lazy val Superscript = TypeOffset("Superscript")
}
case class TypeOffset private(value : String) {

  override def toString = value

}
