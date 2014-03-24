package com.norbitltd.spoiwo.natures.xlsx

import com.norbitltd.spoiwo.model.Measure._
import com.norbitltd.spoiwo.model.{Font, CellStyle, Row}
import org.apache.poi.xssf.usermodel.{XSSFSheet, XSSFFont, XSSFRow, XSSFWorkbook}
import Model2XlsxConversions.convertRow
import org.scalatest.FlatSpec

class Model2XlsxConversionsForRowSpec extends FlatSpec {

  private def sheet : XSSFSheet = new XSSFWorkbook().createSheet()

  private def convert(row : Row) : XSSFRow = convertRow(row, sheet)

  private val defaultRow : XSSFRow = convert(Row.Empty)

  "Row conversion" should "return height of 15 points by default" in {
    assert(defaultRow.getHeightInPoints == 15)
    assert(defaultRow.getHeight == 300)
  }

  it should "return height of 13 point when set explicitly" in {
    val model : Row = Row(height = 13 points)
    val xlsx : XSSFRow = convert(model)
    assert(xlsx.getHeightInPoints == 13)
    assert(xlsx.getHeight == 260)
  }

  it should "return no row style by defaylt" in {
    assert(defaultRow.getRowStyle == null)
  }

  it should "return 14pt Arial row style when set explicitly" in {
    val model : Row = Row(style = CellStyle(font = Font(fontName = "Arial", height = 14 points)))
    val xlsx : XSSFRow = convert(model)

    val font : XSSFFont = xlsx.getRowStyle.getFont
    assert(font.getFontHeightInPoints == 14)
    assert(font.getFontName == "Arial")
  }

  it should "set unhidden row by default" in {
    assert(!defaultRow.getZeroHeight)
  }

  it should "set unhidden row when explicitly set to false" in {
    val model : Row = Row(hidden = false)
    val xlsx : XSSFRow = convert(model)
    assert(!xlsx.getZeroHeight)
  }

  it should "set hidden row when explicitly set to true" in {
    val model : Row = Row(hidden = true)
    val xlsx : XSSFRow = convert(model)
    assert(xlsx.getZeroHeight)
  }

  it should "return row index of 0 for a single row in a sheet by default" in {
    assert(defaultRow.getRowNum == 0)
  }

  it should "return row index of 2 for a single sheet with already 2 rows by default" in {
    val sheetWithRows = sheet
    sheetWithRows.createRow(0)
    sheetWithRows.createRow(1)

    val xlsx : XSSFRow = convertRow(Row(), sheetWithRows)
    assert(xlsx.getRowNum == 2)
  }

}
