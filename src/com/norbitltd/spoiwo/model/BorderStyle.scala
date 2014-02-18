package com.norbitltd.spoiwo.model

object BorderStyle {

  lazy val DashDot = BorderStyle(1)
  lazy val DashDotDot = BorderStyle(2)
  lazy val Dashed = BorderStyle(3)
  lazy val Dotted = BorderStyle(4)
  lazy val Double = BorderStyle(5)
  lazy val Hair = BorderStyle(6)
  lazy val Medium = BorderStyle(7)
  lazy val MediumDashDot = BorderStyle(8)
  lazy val MediumDashDotDot = BorderStyle(9)
  lazy val MediumDashed = BorderStyle(10)
  lazy val None = BorderStyle(11)
  lazy val SlantedDashDot = BorderStyle(12)
  lazy val Thick = BorderStyle(13)
  lazy val Thin = BorderStyle(14) 

}

case class BorderStyle private(id : Int) {

  import BorderStyle._

  def convert() : org.apache.poi.ss.usermodel.BorderStyle = this match {
    case DashDot => org.apache.poi.ss.usermodel.BorderStyle.DASH_DOT
    case DashDotDot => org.apache.poi.ss.usermodel.BorderStyle.DASH_DOT_DOT
    case Dashed => org.apache.poi.ss.usermodel.BorderStyle.DASHED
    case Dotted => org.apache.poi.ss.usermodel.BorderStyle.DOTTED
    case Double => org.apache.poi.ss.usermodel.BorderStyle.DOUBLE
    case Hair => org.apache.poi.ss.usermodel.BorderStyle.HAIR
    case Medium => org.apache.poi.ss.usermodel.BorderStyle.MEDIUM
    case MediumDashDot => org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASH_DOT
    case MediumDashDotDot => org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASH_DOT_DOTC
    case None => org.apache.poi.ss.usermodel.BorderStyle.NONE
    case SlantedDashDot => org.apache.poi.ss.usermodel.BorderStyle.SLANTED_DASH_DOT
    case Thick => org.apache.poi.ss.usermodel.BorderStyle.THICK
    case Thin => org.apache.poi.ss.usermodel.BorderStyle.THIN
  }
}


