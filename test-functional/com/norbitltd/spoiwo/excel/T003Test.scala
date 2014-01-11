package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFDataFormat
import java.util.{Calendar, Date}

class T003Test {

  val format = XSSFDataFormat

  def creatingDateCells() {
    val dateCellStyle = CellStyle(dataFormat = CellDataFormat("m/d/yy h:mm"))
    Sheet(
      name = "new sheet",
      rows = Row(
        Cell(new Date()),
        Cell(new Date(), dateCellStyle),
        Cell(Calendar.getInstance(), dateCellStyle)
      ) :: Nil
    ).save("workbook.xlsx")
  }

}
