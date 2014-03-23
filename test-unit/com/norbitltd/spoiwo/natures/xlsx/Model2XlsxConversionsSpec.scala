package com.norbitltd.spoiwo.natures.xlsx

import org.scalatest.FlatSpec
import com.norbitltd.spoiwo.model._
import Model2XlsxConversions._


class Model2XlsxConversionsSpec extends FlatSpec {

  "Sheet converter" should "set bold font in cell" in {
    val sheetModel = Sheet(name = "TestSheet", row = Row().withCells(Cell("Test", style = CellStyle(font = Font(bold = true)))))
    val converted = sheetModel.convertAsXlsx()

    val convertedSheet = converted.getSheetAt(0)
    val convertedRow = convertedSheet.getRow(0)

    println(convertedRow.getCell(0).getCellStyle.getFont.getBold)
  }
}
