package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFSheet


case class PrintSetup(copies : Short = 1, draft : Boolean = false, fitHeight : Short, fitWidth : Short) {

  def applyTo(sheet : XSSFSheet) = {
    val printSetup = sheet.getPrintSetup
    printSetup.setCopies(copies)
    printSetup.setDraft(draft)
    //TODO
  }
}
