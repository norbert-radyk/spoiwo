package com.norbitltd.spoiwo.model

class T009Test {

  def workingWithFonts() {
    val font = Font(
      height = 24, fontName = "Courier New", italic = true, strikeout = true
    )

    Sheet(Row(Cell("This is a test of fonts", CellStyle(font = font))))
      .withSheetName("new sheet")
      .saveAsXlsx("workbook.xls")
  }
}
