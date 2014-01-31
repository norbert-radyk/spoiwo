package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait Factory {

  lazy val defaultPOIWorkbook = new XSSFWorkbook()
  lazy val defaultPOISheet = defaultPOIWorkbook.createSheet()

  def wrap[T](value : T, defaultValue : T) = if( value != defaultValue) {
    Option(value)
  } else {
    None
  }

}
