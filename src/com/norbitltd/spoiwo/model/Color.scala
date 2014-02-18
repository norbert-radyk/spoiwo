package com.norbitltd.spoiwo.model

import org.apache.poi.xssf.usermodel.XSSFColor

object Color {

  private[model] val Undefined = Color(-1, 0, 0)
  val Black = Color(0, 0, 0)
  val White = Color(255, 255, 255)
  val Red = Color(255, 0, 0)
  val Lime = Color(0, 255, 0)
  val Blue = Color(0, 0, 255)
  val Yellow = Color(255, 255, 0)
  val Aqua = Color(0, 255, 255)
  val Magenta = Color(255, 0, 255)
  val Silver = Color(192, 192, 192)
  val Gray = Color(128, 128, 128)
  val Maroon = Color(128, 0, 0)
  val Olive = Color(128, 128, 0)
  val Green = Color(0, 128, 0)
  val Purple = Color(128, 0, 128)
  val Teal = Color(0, 128, 128)
  val Navy = Color(0, 0, 128)
  val Orange = Color(255, 165, 0)

  def apply(color: XSSFColor): Color = Option(color).map(c => {
    val rgbArray = c.getRgb
    Color(rgbArray(0), rgbArray(1), rgbArray(2))
  }).getOrElse(null)
}

case class Color(r: Int, g: Int, b: Int) {
  require(r >= 0 && r <= 255 || r == -1, "Invalid red color value = %d".format(r))
  require(g >= 0 && g <= 255, "Invalid green color value = %d".format(g))
  require(b >= 0 && b <= 255, "Invalid blue color value = %d".format(b))

  def convert() = new XSSFColor(Array[Byte](r.toByte, g.toByte, b.toByte))
}
