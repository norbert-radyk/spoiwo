package spoiwo.model.enums

object Charset {

  lazy val ANSI: Charset = Charset("ANSI")
  lazy val Arabic: Charset = Charset("Arabic")
  lazy val Baltic: Charset = Charset("Baltic")
  lazy val ChineseBig5: Charset = Charset("ChineseBig5")
  lazy val Default: Charset = Charset("Default")
  lazy val EastEurope: Charset = Charset("EastEurope")
  lazy val GB2312: Charset = Charset("GB2312")
  lazy val Greek: Charset = Charset("Greek")
  lazy val Hangeul: Charset = Charset("Hangeul")
  lazy val Hebrew: Charset = Charset("Hebrew")
  lazy val Johab: Charset = Charset("Johab")
  lazy val Mac: Charset = Charset("Mac")
  lazy val OEM: Charset = Charset("OEM")
  lazy val Russian: Charset = Charset("Russian")
  lazy val ShiftJIS: Charset = Charset("ShiftJIS")
  lazy val Symbol: Charset = Charset("Symbol")
  lazy val Thai: Charset = Charset("Thai")
  lazy val Turkish: Charset = Charset("Turkish")
  lazy val Vietnamese: Charset = Charset("Vietnamese")

}

case class Charset private (value: String) {

  override def toString: String = value

}
