package com.norbitltd.spoiwo.natures.xlsx

import org.scalatest.FlatSpec
import org.apache.poi.xssf.usermodel.{XSSFCellStyle, XSSFWorkbook, XSSFSheet}
import com.norbitltd.spoiwo.model.Column
import Model2XlsxConversions.convertColumn
import com.norbitltd.spoiwo.model.Width._

class Model2XlsxConversionsForColumnSpec extends FlatSpec {

  private def sheet: XSSFSheet = {
    val s = new XSSFWorkbook().createSheet()
    val r = s.createRow(0)
    val c = r.createCell(0)
    c.setCellValue("Test cell")
    s
  }

  private def apply(column: Column): XSSFSheet = {
    val s = sheet
    convertColumn(column, s)
    s
  }

  private def defaultSheet = apply(Column(index = 0))

  "Column conversions" should "have the column width of 8 characters by default" in {
    assert(defaultSheet.getColumnWidth(0) == (8 characters).toUnits)
  }

  it should "have the column width characters of 12 when explicitly specified" in {
    val model = Column(index = 0, width = 12 characters)
    val xlsx = apply(model)
    assert(xlsx.getColumnWidth(0) == (12 characters).toUnits)
  }

  it should "have 11pt Calibri cell style for column by default" in {
    val cellStyle = defaultSheet.getColumnStyle(0).asInstanceOf[XSSFCellStyle]
    assert(cellStyle.getFont.getFontName == "Calibri")
    assert(cellStyle.getFont.getFontHeightInPoints == 11)
  }

  it should "not be hidden by default" in {
    assert(!defaultSheet.isColumnHidden(0))
  }

  it should "not be hidden when explicitly set to false" in {
    val model = Column(index = 0, hidden = false)
    val xlsx = apply(model)
    assert(!xlsx.isColumnHidden(0))
  }

  it should "be hidden when explicitly set to true" in {
    val model = Column(index = 0, hidden = true)
    val xlsx = apply(model)
    assert(xlsx.isColumnHidden(0))
  }

  it should "not be a break by default" in {
    assert(defaultSheet.getColumnBreaks.isEmpty)
  }

  it should "not be a break when set to false" in {
    val model = Column(index = 2, break = false)
    val xlsx = apply(model)
    assert(xlsx.getColumnBreaks.isEmpty)
  }

  it should "be a break when set to true" in {
    val model = Column(index = 2, break = true)
    val xlsx = apply(model)
    assert(xlsx.getColumnBreaks.toSet == Set(2))
  }

  it should "remain in default width when auto-sized set to false" in {
    val model = Column(index = 0, autoSized = false)
    val xlsx = apply(model)
    assert(xlsx.getColumnWidth(0) == (8 characters).toUnits)
  }

  it should "resize column width to the content when auto-sized set to true" in {
    val model = Column(index = 0, autoSized = true)
    val xlsx = apply(model)
    assert(xlsx.getColumnWidth(0) == 2129)
  }

  //TODO Implement testing for group collapsed

}
