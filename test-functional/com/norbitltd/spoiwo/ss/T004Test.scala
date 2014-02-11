package com.norbitltd.spoiwo.ss

import java.util.{Date, Calendar}

class T004Test {

  val sheet = Sheet(
    Row.Empty,
    Row.Empty,
    Row(Cell(1.1), Cell(new Date()), Cell(Calendar.getInstance()), Cell("a string"), Cell("=1/0"))
  ).withSheetName("new sheet")

  sheet.saveAsXlsx("workbook.xlsx")
}
