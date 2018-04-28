package com.norbitltd.spoiwo.model.enums

object FontScheme {
  lazy val None = FontScheme("None")
  lazy val Major = FontScheme("Major")
  lazy val Minor = FontScheme("Minor")
}

case class FontScheme private (value: String) {

  override def toString: String = value

}
