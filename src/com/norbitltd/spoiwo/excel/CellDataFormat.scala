package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.{XSSFDataFormat, XSSFWorkbook, XSSFCellStyle}

object CellDataFormat {

  val Undefined = apply(None)

  def apply(formatString : String) = CellDataFormat(Option(formatString))

  val cache = collection.mutable.Map[XSSFWorkbook, XSSFDataFormat]()

}

case class CellDataFormat private[excel](formatString : Option[String]) {

  private def convert(workbook : XSSFWorkbook) : Short = {
    val dataFormat = CellDataFormat.cache.getOrElseUpdate(workbook, workbook.createDataFormat());
    dataFormat.getFormat(formatString.get)
  }

  def applyTo(workbook : XSSFWorkbook, cellStyle : XSSFCellStyle) {
    if( formatString.isDefined ) {
      cellStyle.setDataFormat(convert(workbook))
    }
  }

}
