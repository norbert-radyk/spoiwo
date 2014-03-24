package com.norbitltd.spoiwo.natures.xlsx

import org.scalatest.FlatSpec
import org.apache.poi.xssf.usermodel.{XSSFWorkbook, XSSFSheet}
import com.norbitltd.spoiwo.model.Column
import Model2XlsxConversions.convertColumn

class Model2XlsxConversionsForColumnSpec extends FlatSpec {

  private def sheet : XSSFSheet = new XSSFWorkbook().createSheet()

  private def apply(column : Column) : XSSFSheet = {
    val s = sheet
    convertColumn(column, s)
    s
  }

  private def defaultSheet = apply(Column())



}
