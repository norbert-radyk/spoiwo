package com.norbitltd.spoiwo.model.enums

object Charset {

  lazy val ANSI = Charset("ANSI")
  lazy val Arabic = Charset("Arabic")
  lazy val Baltic = Charset("Baltic")
  lazy val ChineseBig5 = Charset("ChineseBig5")
  lazy val Default = Charset("Default")
  lazy val EastEurope = Charset("EastEurope")
  lazy val GB2312 = Charset("GB2312")
  lazy val Greek = Charset("Greek")
  lazy val Hangeul = Charset("Hangeul")
  lazy val Hebrew = Charset("Hebrew")
  lazy val Johab = Charset("Johab")
  lazy val Mac = Charset("Mac")
  lazy val OEM = Charset("OEM")
  lazy val Russian = Charset("Russian")
  lazy val ShiftJIS = Charset("ShiftJIS")
  lazy val Symbol = Charset("Symbol")
  lazy val Thai = Charset("Thai")
  lazy val Turkish = Charset("Turkish")
  lazy val Vietnamese = Charset("Vietnamese")

}

case class Charset private (value: String) {

  override def toString: String = value

}
