package com.norbitltd.spoiwo.model.enums

object FontFamily {

  private[model] lazy val Undefined = FontFamily("Undefined")
  lazy val NotApplicable = FontFamily("NotApplicable")
  lazy val Roman = FontFamily("Roman")
  lazy val Swiss = FontFamily("Swiss")
  lazy val Modern = FontFamily("Modern")
  lazy val Script = FontFamily("Script")
  lazy val Decorative = FontFamily("Decorative")

}

case class FontFamily private(value : String) {

  override def toString = value

}
