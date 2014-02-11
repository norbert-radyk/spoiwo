package com.norbitltd.spoiwo.ss

import org.apache.poi.ss.usermodel.FillPatternType

class T007Test {

  def fillsAndColors() {

    val contentRow = Row(
      Cell("X", CellStyle(fillBackgroundColor = Color.AQUA, fillPattern = FillPatternType.BIG_SPOTS)),
      Cell("X", CellStyle(fillForegroundColor = Color.ORANGE, fillPattern = FillPatternType.SOLID_FOREGROUND))
    )

    Sheet(name = "new sheet", rows = contentRow :: Nil)
      .saveAsXlsx("workbook.xlsx")
  }

}
