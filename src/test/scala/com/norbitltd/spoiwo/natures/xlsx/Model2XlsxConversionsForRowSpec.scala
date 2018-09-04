package com.norbitltd.spoiwo.natures.xlsx

import com.norbitltd.spoiwo.model.Height._
import com.norbitltd.spoiwo.model._
import org.apache.poi.xssf.usermodel.{XSSFFont, XSSFRow, XSSFSheet, XSSFWorkbook}
import Model2XlsxConversions.convertRow
import org.scalatest.{FlatSpec, Matchers}

import scala.util.Try
import scala.language.postfixOps

class Model2XlsxConversionsForRowSpec extends FlatSpec with Matchers {

  private def sheet: XSSFSheet = new XSSFWorkbook().createSheet()

  private def convert(row: Row): XSSFRow = convertRow(Map(), row, Sheet(), sheet)

  private val defaultRow: XSSFRow = convert(Row.Empty)

  "Row conversion" should "return height of 15 points by default" in {
    defaultRow.getHeightInPoints shouldBe 15
    defaultRow.getHeight shouldBe 300
  }

  it should "return height of 13 point when set explicitly" in {
    val model: Row = Row(height = 13 points)
    val xlsx: XSSFRow = convert(model)
    xlsx.getHeightInPoints shouldBe 13
    xlsx.getHeight shouldBe 260
  }

  it should "return no row style by default" in {
    defaultRow.getRowStyle shouldBe null
  }

  it should "return 14pt Arial row style when set explicitly" in {
    val model: Row = Row(style = CellStyle(font = Font(fontName = "Arial", height = 14 points)))
    val xlsx: XSSFRow = convert(model)

    val font: XSSFFont = xlsx.getRowStyle.getFont
    font.getFontHeightInPoints shouldBe 14
    font.getFontName shouldBe "Arial"
  }

  it should "set unhidden row by default" in {
    defaultRow.getZeroHeight shouldBe false
  }

  it should "set unhidden row when explicitly set to false" in {
    val model: Row = Row(hidden = false)
    val xlsx: XSSFRow = convert(model)
    xlsx.getZeroHeight shouldBe false
  }

  it should "set hidden row when explicitly set to true" in {
    val model: Row = Row(hidden = true)
    val xlsx: XSSFRow = convert(model)
    xlsx.getZeroHeight shouldBe true
  }

  it should "return row index of 0 for a single row in a sheet by default" in {
    defaultRow.getRowNum shouldBe 0
  }

  it should "return row index of 2 for a single sheet with already 2 rows by default" in {
    val sheetWithRows = sheet
    sheetWithRows.createRow(0)
    sheetWithRows.createRow(1)

    val xlsx: XSSFRow = convertRow(Map(), Row(), Sheet(), sheetWithRows)
    xlsx.getRowNum shouldBe 2
  }

  it should "return index of 5 if the last row in the sheet has index of 4" in {
    val sheetWithRows = sheet
    sheetWithRows.createRow(4)

    val xlsx: XSSFRow = convertRow(Map(), Row(), Sheet(), sheetWithRows)
    xlsx.getRowNum shouldBe 5
  }

  it should "have no cells by default" in {
    defaultRow.cellIterator().hasNext shouldBe false
  }

  it should "correctly initialize one cell" in {
    val cell: Cell = Cell("Test", style = CellStyle(font = Font(fontName = "Arial", height = 14 points)))
    val model: Row = Row(cells = cell :: Nil)
    val xssf: XSSFRow = convert(model)

    xssf.getFirstCellNum shouldBe 0
    xssf.getLastCellNum shouldBe 1

    val convertedFont = xssf.getCell(0).getCellStyle.getFont
    convertedFont.getFontName shouldBe "Arial"
    convertedFont.getFontHeightInPoints shouldBe 14
  }

  it should "correctly initialize multiple cells with no indexes" in {
    val cell1: Cell = Cell("Test1", style = CellStyle(font = Font(fontName = "Arial", height = 14 points)))
    val cell2: Cell = Cell("Test2", style = CellStyle(font = Font(fontName = "Arial", height = 16 points)))
    val model: Row = Row(cells = cell1 :: cell2 :: Nil)
    val xssf: XSSFRow = convert(model)

    xssf.getFirstCellNum shouldBe 0
    xssf.getLastCellNum shouldBe 2

    val convertedFont1 = xssf.getCell(0).getCellStyle.getFont
    convertedFont1.getFontName shouldBe "Arial"
    convertedFont1.getFontHeightInPoints shouldBe 14

    val convertedFont2 = xssf.getCell(1).getCellStyle.getFont
    convertedFont2.getFontName shouldBe "Arial"
    convertedFont2.getFontHeightInPoints shouldBe 16
  }

  it should "correctly initialize single cell with index" in {
    val cell = Cell("Test", index = 3)
    val model: Row = Row(cells = cell :: Nil)
    val xssf: XSSFRow = convert(model)
    xssf.getFirstCellNum shouldBe 3
    xssf.getLastCellNum shouldBe 4

    val xssfCell = xssf.getCell(3)
    xssfCell.getStringCellValue shouldBe "Test"
  }

  it should "correctly initialize multiple cells with indexes" in {
    val cell1 = Cell("Test1", index = 3)
    val cell2 = Cell("Test2", index = 5)
    val model: Row = Row(cells = cell1 :: cell2 :: Nil)
    val xssf: XSSFRow = convert(model)
    xssf.getFirstCellNum shouldBe 3
    xssf.getLastCellNum shouldBe 6

    val xssfCell1 = xssf.getCell(3)
    xssfCell1.getStringCellValue shouldBe "Test1"

    val xssfCell2 = xssf.getCell(5)
    xssfCell2.getStringCellValue shouldBe "Test2"
  }

  it should "fail when row has 2 cells with duplicate indexes" in {
    val cell1 = Cell("Test1", index = 2)
    val cell2 = Cell("Test2", index = 4)
    val cell3 = Cell("Test3", index = 2)
    val model: Row = Row(cells = cell1 :: cell2 :: cell3 :: Nil)

    val converted = Try {

      convert(model)
    }
    converted.isFailure shouldBe true
  }

  it should "fail when row has cells with and without explicit indexes" in {
    val cell1 = Cell("Test1", index = 2)
    val cell2 = Cell("Test2")
    val model: Row = Row(cells = cell1 :: cell2 :: Nil)

    val converted = Try {

      convert(model)
    }
    converted.isFailure shouldBe true
  }
}
