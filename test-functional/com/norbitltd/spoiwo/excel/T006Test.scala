package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.usermodel.BorderStyle._
import com.norbitltd.spoiwo.excel.Color._


class T006Test {

  def workingWithBorders() {

    val borderStyle = CellStyle(borders = CellBorder(
        bottomStyle = THIN, bottomColor = BLACK,
        leftStyle = THIN, leftColor = GREEN,
        rightStyle = THIN, rightColor = BLUE,
        topStyle = MEDIUM_DASHED, topColor = BLACK))

    Sheet( Row(Cell(4, borderStyle)) )
      .withSheetName("new sheet")
      .save("workbook.xlsx")
  }
}
