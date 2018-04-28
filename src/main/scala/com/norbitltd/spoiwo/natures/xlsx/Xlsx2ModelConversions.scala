package com.norbitltd.spoiwo.natures.xlsx

import org.apache.poi.xssf.usermodel.XSSFColor
import com.norbitltd.spoiwo.model.Color

object Xlsx2ModelConversions {

  def convertXSSFColorSafe(color: XSSFColor): Option[Color] = {
    Option(color).map(convertXSSFColor)
  }

  def convertXSSFColor(color: XSSFColor): Color = {
    require(color != null, "Can't convert NULL color value!")
    Color(color.getRGB)
  }

}
