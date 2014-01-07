package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream


object Workbook {

   def apply(sheets : Sheet*) : Workbook = new Workbook(sheets.toList)

}

class Workbook(sheets : List[Sheet], activeSheet : Int = 0) {
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
