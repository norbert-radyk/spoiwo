package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFColor

object Color {
  val BLACK = Color(0, 0, 0)

  val WHITE = Color(255, 255, 255)
}

case class Color(r : Int, g : Int, b : Int) {
  require(r >= 0 && r <= 255, "Invalid red color value = %d".format(r))
  require(g >= 0 && g <= 255, "Invalid green color value = %d".format(g))
  require(b >= 0 && b <= 255, "Invalid blue color value = %d".format(b))

  def convert() = new XSSFColor(Array[Byte](r.toByte, g.toByte, b.toByte))
}
