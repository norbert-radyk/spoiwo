package com.norbitltd.spoiwo.ss

import org.apache.poi.ss.usermodel.FillPatternType

class T010Test {

  Sheet(Row(
    Cell("custom XSSF colors", CellStyle(
      fillForegroundColor = Color(128, 0, 128),
      fillPattern = FillPatternType.SOLID_FOREGROUND)
    )
  ))

}
