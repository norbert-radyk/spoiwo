package com.norbitltd.spoiwo.model

import org.apache.poi.xssf.usermodel.XSSFColor

object Color {
  val BLACK = Color(0, 0, 0)
  val WHITE = Color(255, 255, 255)
  val RED = Color(255, 0, 0)
  val LIME = Color(0, 255, 0)
  val BLUE = Color(0, 0, 255)
  val YELLOW = Color(255, 255, 0)
  val AQUA = Color(0, 255, 255)
  val MAGENTA = Color(255, 0, 255)
  val SILVER = Color(192, 192, 192)
  val GRAY = Color(128, 128, 128)
  val MAROON = Color(128, 0, 0)
  val OLIVE = Color(128, 128, 0)
  val GREEN = Color(0, 128, 0)
  val PURPLE = Color(128, 0, 128)
  val TEAL = Color(0, 128, 128)
  val NAVY = Color(0, 0, 128)
  val ORANGE = Color(255, 165, 0)

  def apply(color: XSSFColor): Color = Option(color).map(c => {
    val rgbArray = c.getRgb
    Color(rgbArray(0), rgbArray(1), rgbArray(2))
  }).getOrElse(null)
}

case class Color(r: Int, g: Int, b: Int) {
  require(r >= 0 && r <= 255, "Invalid red color value = %d".format(r))
  require(g >= 0 && g <= 255, "Invalid green color value = %d".format(g))
  require(b >= 0 && b <= 255, "Invalid blue color value = %d".format(b))

  def convert() = new XSSFColor(Array[Byte](r.toByte, g.toByte, b.toByte))
}
