package com.norbitltd.spoiwo.model.enums

object FontScheme {
  private[model] lazy val Undefined = FontScheme("Undefined")
  lazy val None = FontScheme("None")
  lazy val Major = FontScheme("Major")
  lazy val Minor = FontScheme("Minor")
}

case class FontScheme private(value: String) {

  override def toString = value

}
