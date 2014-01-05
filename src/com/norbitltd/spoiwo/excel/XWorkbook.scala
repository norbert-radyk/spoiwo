package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream


object XWorkbook {

   def apply(sheets : XSheet*) : XWorkbook = new XWorkbook(sheets.toList)

}

class XWorkbook(sheets : List[XSheet], activeSheet : Int = 0) {
  require(!sheets.isEmpty, "Unable to construct Excel workbook with no sheets!")

  def convert() : XSSFWorkbook = {
    val workbook = new XSSFWorkbook()
    sheets.foreach(sheet => sheet.convert(workbook))
    workbook.setActiveSheet(activeSheet)
    workbook
  }

  def save(fileName : String) = {
    val stream = new FileOutputStream(fileName)
    try {
      val xssfWorkbook = convert()
      xssfWorkbook.write(stream)
    } finally {
      stream.close()
    }
  }
}
