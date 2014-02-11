package com.norbitltd.spoiwo.ss

import org.apache.poi.ss.usermodel.BorderStyle._
import com.norbitltd.spoiwo.ss.Color._


class T006Test {

  def workingWithBorders() {

    val borderStyle = CellStyle(borders = CellBorders(
        bottomStyle = THIN, bottomColor = BLACK,
        leftStyle = THIN, leftColor = GREEN,
        rightStyle = THIN, rightColor = BLUE,
        topStyle = MEDIUM_DASHED, topColor = BLACK))

    Sheet( Row(Cell(4, borderStyle)) )
      .withSheetName("new sheet")
      .saveAsXlsx("workbook.xlsx")
  }
}
