package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.usermodel.{VerticalAlignment, HorizontalAlignment}

class T005Test {

  def createCell(halign : HorizontalAlignment, valign : VerticalAlignment) =
    Cell("Align It", CellStyle(horizontalAlignment = halign, verticalAlignment = valign))

  def aligningCells() {
    val rowWithAlignments = Row(
      createCell(HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM),
      createCell(HorizontalAlignment.CENTER_SELECTION, VerticalAlignment.BOTTOM),
      createCell(HorizontalAlignment.FILL, VerticalAlignment.CENTER),
      createCell(HorizontalAlignment.GENERAL, VerticalAlignment.CENTER),
      createCell(HorizontalAlignment.JUSTIFY, VerticalAlignment.JUSTIFY),
      createCell(HorizontalAlignment.LEFT, VerticalAlignment.TOP),
      createCell(HorizontalAlignment.RIGHT, VerticalAlignment.TOP)
    ).withHeightInPoints(10)

    Sheet(Row.Empty, Row.Empty, rowWithAlignments).save("xssf-align.xlsx")
  }

}
