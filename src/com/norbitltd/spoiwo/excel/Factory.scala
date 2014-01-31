package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait Factory {

  lazy val defaultPOIWorkbook = new XSSFWorkbook()
  lazy val defaultPOISheet = defaultPOIWorkbook.createSheet()
  lazy val defaultPOICellStyle = defaultPOIWorkbook.createCellStyle()
  lazy val defaultPOIFont = defaultPOIWorkbook.createFont()

  def wrap[T](value : T, defaultValue : T) : Option[T] = if( value != defaultValue) {
    Option(value)
  } else {
    None
  }

}
