package com.norbitltd.spoiwo.ss

import org.apache.poi.ss.util.WorkbookUtil

class T001Test {

  def generateWorkbookAndSheet() {
    Workbook(
      Sheet(name = "new sheet"),
      Sheet(name = "second sheet"),
      Sheet(name = WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]"))
    ).saveXLSX("workbook.xlsx")
  }


  def test() {
    Row(Cell("AAA"), Cell(2.30))

  }
}
