package com.norbitltd.spoiwo.model

import org.apache.poi.xssf.usermodel.XSSFWorkbook

trait Factory {

  lazy val defaultPOIWorkbook = new XSSFWorkbook()
  lazy val defaultPOISheet = defaultPOIWorkbook.createSheet()
  lazy val defaultPOIPrintSetup = defaultPOISheet.getPrintSetup

  def wrap[T](value: T, defaultValue: T): Option[T] = if (value != defaultValue) {
    Option(value)
  } else {
    None
  }

}
