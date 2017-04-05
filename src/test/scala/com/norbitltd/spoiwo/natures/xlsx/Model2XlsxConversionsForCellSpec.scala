package com.norbitltd.spoiwo.natures.xlsx

import java.util.Calendar
import java.time.{LocalDate => JLocalDate, LocalDateTime => JLocalDateTime}
import com.norbitltd.spoiwo.model.Height._
import com.norbitltd.spoiwo.model._
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions.{convertCell, _}
import org.apache.poi.ss.usermodel
import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFWorkbook}
import org.joda.time.{DateTime, LocalDate}
import org.scalatest.FlatSpec
import scala.language.postfixOps


class Model2XlsxConversionsForCellSpec extends FlatSpec {

  private val defaultCell: XSSFCell = convert(Cell.Empty)

  private def convert(cell: Cell): XSSFCell = convertCell(Sheet(), Map(), Row(), cell, row)

  private def row = new XSSFWorkbook().createSheet().createRow(0)

  "Cell conversion" should "return string cell type with empty string by default" in {
    assert(defaultCell.getCellType == usermodel.Cell.CELL_TYPE_STRING)
    assert(defaultCell.getStringCellValue == "")
  }

  it should "return a 0 column index by default if no other cells specified in the row" in {
    assert(defaultCell.getColumnIndex == 0)
  }

  it should "return cell style with 11pt Calibri by default" in {
    assert(defaultCell.getCellStyle.getFont.getFontHeightInPoints === 11)
    assert(defaultCell.getCellStyle.getFont.getFontName === "Calibri")
  }

  it should "return cell style with 14pt Arial when explicitly specified" in {
    val cellStyle = CellStyle(font = Font(fontName = "Arial", height = 14 points))
    val model = Cell.Empty.withStyle(cellStyle)
    val xlsx = convert(model)

    assert(xlsx.getCellStyle.getFont.getFontHeightInPoints == 14)
    assert(xlsx.getCellStyle.getFont.getFontName == "Arial")
  }

  it should "return index of 3 when explicitly specified" in {
    val model = Cell.Empty.withIndex(3)
    val xlsx = convert(model)

    assert(xlsx.getColumnIndex == 3)
  }

  it should "return index of 2 when row has already 2 other cells" in {
    val row = new XSSFWorkbook().createSheet().createRow(0)
    row.createCell(0)
    row.createCell(1)

    val model = Cell.Empty
    val xlsx = convertCell(Sheet(), Map(), Row(), model, row)
    assert(xlsx.getColumnIndex == 2)
  }

  it should "return string cell when set up with 'String'" in {
    val model = Cell("TEST_STRING")
    val xlsx = convert(model)
    assert(xlsx.getCellType == usermodel.Cell.CELL_TYPE_STRING)
    assert(xlsx.getStringCellValue == "TEST_STRING")
  }

  it should "return string cell when set up with String with newline value" in {
    val model = Cell("TEST_STRING\nAnd a 2nd line")
    val xlsx = convert(model)
    assert(xlsx.getCellType == usermodel.Cell.CELL_TYPE_STRING)
    assert(xlsx.getStringCellValue == "TEST_STRING\nAnd a 2nd line")
  }

  it should "return formula cell when set up with string starting with '=' sign" in {
    val model = Cell("=1000/3+7")
    val xlsx = convert(model)
    assert(xlsx.getCellType == usermodel.Cell.CELL_TYPE_FORMULA)
    assert(xlsx.getCellFormula == "1000/3+7")
  }

  it should "return numeric cell when set up with double value" in {
    val model = Cell(90.45)
    val xlsx = convert(model)
    assert(xlsx.getCellType == usermodel.Cell.CELL_TYPE_NUMERIC)
    assert(xlsx.getNumericCellValue == 90.45)
  }

  it should "return numeric cell when set up with big decimal value" in {
    val model = Cell(BigDecimal(90.45))
    val xlsx = convert(model)
    assert(xlsx.getCellType == usermodel.Cell.CELL_TYPE_NUMERIC)
    assert(xlsx.getNumericCellValue == 90.45)
  }

  it should "return numeric cell when set up with int value" in {
    val model = Cell(90)
    val xlsx = convert(model)
    assert(xlsx.getCellType == usermodel.Cell.CELL_TYPE_NUMERIC)
    assert(xlsx.getNumericCellValue == 90)
  }

  it should "return numeric cell when set up with long value" in {
    val model = Cell(10000000000000l)
    val xlsx = convert(model)
    assert(xlsx.getCellType == usermodel.Cell.CELL_TYPE_NUMERIC)
    assert(xlsx.getNumericCellValue == 10000000000000l)
  }

  it should "return boolean cell when set up with boolean value" in {
    val model = Cell(true)
    val xlsx = convert(model)
    assert(xlsx.getCellType == usermodel.Cell.CELL_TYPE_BOOLEAN)
    assert(xlsx.getBooleanCellValue)
  }

  it should "return numeric cell when set up with java.util.Date value" in {
    val model = Cell(new LocalDate(2011, 11, 13).toDate)
    val xlsx = convert(model)

    val date = new LocalDate(xlsx.getDateCellValue)
    assert(date.getYear == 2011)
    assert(date.getMonthOfYear == 11)
    assert(date.getDayOfMonth == 13)
  }

  it should "return numeric cell when set up with org.joda.time.LocalDate value" in {
    val model = Cell(new LocalDate(2011, 11, 13))
    val xlsx = convert(model)

    val date = new LocalDate(xlsx.getDateCellValue)
    assert(date.getYear == 2011)
    assert(date.getMonthOfYear == 11)
    assert(date.getDayOfMonth == 13)
  }

  it should "return numeric cell when set up with java.util.Calendar value" in {
    val calendar = Calendar.getInstance()
    calendar.set(2011, 11, 13)
    val model = Cell(calendar)
    val xlsx = convert(model)

    val date = new LocalDate(xlsx.getDateCellValue)
    assert(date.getYear == 2011)
    assert(date.getMonthOfYear == 12)
    assert(date.getDayOfMonth == 13)
  }

  it should "return numeric cell when set up with java.time.LocalDate value" in {
    test(JLocalDate.of(2011, 6, 13))
    test(JLocalDate.of(2011, 11, 13))

    def test(ld: JLocalDate): Unit = {
      val model = Cell(ld)
      val xlsx = convert(model)

      val date = new DateTime(xlsx.getDateCellValue)
      assert(date.getYear           == ld.getYear)
      assert(date.getMonthOfYear    == ld.getMonthValue)
      assert(date.getDayOfMonth     == ld.getDayOfMonth)
      assert(date.getHourOfDay      == 0)
      assert(date.getMinuteOfHour   == 0)
      assert(date.getSecondOfMinute == 0)
    }
  }

  it should "return numeric cell when set up with java.time.LocalDateTime value" in {
    test(JLocalDateTime.of(2011,  6, 13, 15, 30, 10))
    test(JLocalDateTime.of(2011, 11, 13, 15, 30, 10))

    def test(ldt: JLocalDateTime): Unit = {
      val model = Cell(ldt)
      val xlsx = convert(model)

      val date = new DateTime(xlsx.getDateCellValue)
      assert(date.getYear           == ldt.getYear)
      assert(date.getMonthOfYear    == ldt.getMonthValue)
      assert(date.getDayOfMonth     == ldt.getDayOfMonth)
      assert(date.getHourOfDay      == ldt.getHour)
      assert(date.getMinuteOfHour   == ldt.getMinute)
      assert(date.getSecondOfMinute == ldt.getSecond)
    }
  }

  it should "return string cell with the date formatted yyyy-MM-dd if date before 1904" in {
    val model = Cell(new LocalDate(1856, 11, 3).toDate)
    val xlsx = convert(model)
    assert("1856-11-03" === xlsx.getStringCellValue)
  }

  it should "apply 14pt Arial cell style for column when set explicitly" in {
    val column = Column(index = 0, style = CellStyle(font = Font(fontName = "Arial", height = 14 points)))
    val sheet = Sheet(Row(Cell("Test"))).withColumns(column)
    val xlsx = sheet.convertAsXlsx()

    val cellStyle = xlsx.getSheetAt(0).getRow(0).getCell(0).getCellStyle
    assert(cellStyle.getFont.getFontName == "Arial")
    assert(cellStyle.getFont.getFontHeightInPoints == 14)
  }

  it should "return a string cell with hyperlink when setup with HyperLinkUrl value" in {
    val model = Cell(HyperLinkUrl("View Item", "https://www.google.com"))
    val xlsx = convert(model)
    assert(xlsx.getCellType == usermodel.Cell.CELL_TYPE_STRING)
    assert(xlsx.getHyperlink().getAddress() == "https://www.google.com")
  }
}
