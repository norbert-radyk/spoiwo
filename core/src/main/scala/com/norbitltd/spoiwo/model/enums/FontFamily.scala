package com.norbitltd.spoiwo.model.enums

object FontFamily {

  lazy val NotApplicable: FontFamily = FontFamily("NotApplicable")
  lazy val Roman: FontFamily = FontFamily("Roman")
  lazy val Swiss: FontFamily = FontFamily("Swiss")
  lazy val Modern: FontFamily = FontFamily("Modern")
  lazy val Script: FontFamily = FontFamily("Script")
  lazy val Decorative: FontFamily = FontFamily("Decorative")

}

case class FontFamily private (value: String) {

  override def toString: String = value

}
