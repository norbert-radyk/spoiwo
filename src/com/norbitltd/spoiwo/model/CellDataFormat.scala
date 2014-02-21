package com.norbitltd.spoiwo.model

import org.apache.poi.xssf.usermodel.{XSSFDataFormat, XSSFWorkbook, XSSFCellStyle}

object CellDataFormat {

  val Undefined = apply(None)

  def apply(formatString : String) : CellDataFormat = CellDataFormat(Option(formatString))

  val cache = collection.mutable.Map[XSSFWorkbook, XSSFDataFormat]()

}

case class CellDataFormat private[model](formatString : Option[String]) {

  private def convert(workbook : XSSFWorkbook) : Short = {
    val dataFormat = CellDataFormat.cache.getOrElseUpdate(workbook, workbook.createDataFormat())
    dataFormat.getFormat(formatString.get)
  }

  def applyTo(workbook : XSSFWorkbook, cellStyle : XSSFCellStyle) {
    if( formatString.isDefined ) {
      cellStyle.setDataFormat(convert(workbook))
    }
  }

}
