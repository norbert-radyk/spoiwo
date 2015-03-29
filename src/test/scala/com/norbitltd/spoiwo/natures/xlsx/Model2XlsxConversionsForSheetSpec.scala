package com.norbitltd.spoiwo.natures.xlsx

import org.scalatest.FlatSpec
import org.apache.poi.xssf.usermodel.{XSSFCellStyle, XSSFSheet, XSSFWorkbook}
import Model2XlsxConversions.convertSheet
import com.norbitltd.spoiwo.model._
import com.norbitltd.spoiwo.model.Height._
import com.norbitltd.spoiwo.model.Width._
import scala.util.Try
import scala.language.postfixOps

class Model2XlsxConversionsForSheetSpec extends FlatSpec {

  private def workbook: XSSFWorkbook = new XSSFWorkbook()

  private def convert(s: Sheet): XSSFSheet = convertSheet(s, workbook)

  private def defaultSheet = convert(Sheet())

  "Sheet conversions" should "set sheet name to 'Sheet1' by default in empty workbook" in {
    assert(defaultSheet.getSheetName == "Sheet1")
  }

  it should "set sheet name to 'Sheet 3'in a workbook with already 2 sheets defined" in {
    val w: XSSFWorkbook = workbook
    w.createSheet("Test1")
    w.createSheet("Test2")
    val sheet3: XSSFSheet = convertSheet(Sheet(), w)
    assert(sheet3.getSheetName == "Sheet3")
  }

  it should "return defined name when set explicitly" in {
    val model = Sheet(name = "My Name")
    val xssf = convert(model)
    assert(xssf.getSheetName === "My Name")
  }

  it should "not have 1st column style by default" in {
    val xssfFont = defaultSheet.getColumnStyle(0).asInstanceOf[XSSFCellStyle].getFont
    assert(xssfFont.getFontName == "Calibri")
    assert(xssfFont.getFontHeightInPoints == 11)
  }

  it should "have 1st column width when set with column without index" in {
    val model = Sheet(columns = Column(width = 200 unitsOfWidth) :: Nil)
    val xssf = convert(model)
    assert(xssf.getColumnWidth(0) == 200)
  }

  it should "have correct 1st and 2nd column style when set with column without index" in {
    val column1 = Column(width = 200 unitsOfWidth)
    val column2 = Column(width = 300 unitsOfWidth)
    val model = Sheet(columns = column1 :: column2 :: Nil)
    val xssf = convert(model)

    assert(xssf.getColumnWidth(0) == 200)
    assert(xssf.getColumnWidth(1) == 300)
  }

  it should "have 3rd column style when set with column with index" in {
    val model = Sheet(columns = Column(width = 300 unitsOfWidth, index = 2) :: Nil)
    val xssf = convert(model)
    assert(xssf.getColumnWidth(2) == 300)
  }

  it should "have correct 4th and 6th column style when set with column without index" in {
    val column1 = Column(width = 200 unitsOfWidth, index = 3)
    val column2 = Column(width = 300 unitsOfWidth, index = 5)
    val model = Sheet(columns = column1 :: column2 :: Nil)
    val xssf = convert(model)

    assert(xssf.getColumnWidth(3) == 200)
    assert(xssf.getColumnWidth(4) == (8 characters).toUnits)
    assert(xssf.getColumnWidth(5) == 300)
  }

  it should "not allow mixing columns with and without index" in {
    val column1 = Column(style = CellStyle(font = Font(fontName = "Arial", height = 14 points)))
    val column2 = Column(style = CellStyle(font = Font(fontName = "Arial", height = 16 points)), index = 5)
    val model = Sheet(columns = column1 :: column2 :: Nil)
    val xlsx = Try(convert(model))
    assert(xlsx.isFailure)
  }

  it should "not allow 2 columns with the same index" in {
    val column1 = Column(style = CellStyle(font = Font(fontName = "Arial", height = 14 points)), index = 3)
    val column2 = Column(style = CellStyle(font = Font(fontName = "Arial", height = 16 points)), index = 5)
    val column3 = Column(style = CellStyle(font = Font(fontName = "Arial", height = 16 points)), index = 3)
    val model = Sheet(columns = column1 :: column2 :: column3 :: Nil)
    val xlsx = Try(convert(model))
    assert(xlsx.isFailure)
  }

  it should "have no rows by default" in {
    assert(!defaultSheet.rowIterator().hasNext)
  }

  it should "have 1st row with initialized with row and without index" in {
    val row = Row().withCellValues("Test", 34)
    val model = Sheet(rows = row :: Nil)

    val xlsx = convert(model)
    val xlsxRow = xlsx.getRow(0)
    assert(xlsxRow.getCell(0).getStringCellValue == "Test")
    assert(xlsxRow.getCell(1).getNumericCellValue == 34)
  }

  it should "have 1st and 2nd row with initialized with row and without index" in {
    val row1 = Row().withCellValues("Test", 34)
    val row2 = Row().withCellValues("OK")
    val model = Sheet(rows = row1 :: row2 :: Nil)

    val xlsx = convert(model)
    val xlsxRow1 = xlsx.getRow(0)
    val xlsxRow2 = xlsx.getRow(1)
    assert(xlsxRow1.getCell(0).getStringCellValue == "Test")
    assert(xlsxRow1.getCell(1).getNumericCellValue == 34)
    assert(xlsxRow2.getCell(0).getStringCellValue == "OK")
  }


  it should "have 3rd row with initialized with row with index" in {
    val row = Row(index = 2).withCellValues("Test", 34)
    val model = Sheet(rows = row :: Nil)

    val xlsx = convert(model)
    val xlsxRow = xlsx.getRow(2)
    assert(xlsxRow.getCell(0).getStringCellValue == "Test")
    assert(xlsxRow.getCell(1).getNumericCellValue == 34)
  }

  it should "have 3rd and 5th row with initialized with row and index" in {
    val row1 = Row(index = 4).withCellValues("Test", 34)
    val row2 = Row(index = 6).withCellValues("OK")
    val model = Sheet(rows = row1 :: row2 :: Nil)

    val xlsx = convert(model)
    val xlsxRow1 = xlsx.getRow(4)
    val xlsxRow2 = xlsx.getRow(6)
    assert(xlsxRow1.getCell(0).getStringCellValue == "Test")
    assert(xlsxRow1.getCell(1).getNumericCellValue == 34)
    assert(xlsxRow2.getCell(0).getStringCellValue == "OK")
  }

  it should "not allow duplicate row indexes" in {
    val row1 = Row(index = 4).withCellValues("Test", 34)
    val row2 = Row(index = 6).withCellValues("OK")
    val row3 = Row(index = 4).withCellValues("Override")
    val model = Sheet(rows = row1 :: row2 :: row3 :: Nil)
    val xlsx = Try(convert(model))
    assert(xlsx.isFailure)
  }

  it should "not allow mixing indexed and non indexed rows" in {
    val row1 = Row(index = 4).withCellValues("Test", 34)
    val row2 = Row().withCellValues("OK")
    val model = Sheet(rows = row1 :: row2 :: Nil)
    val xlsx = Try(convert(model))
    assert(xlsx.isFailure)
  }

}
