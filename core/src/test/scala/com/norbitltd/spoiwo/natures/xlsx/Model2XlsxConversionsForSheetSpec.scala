package com.norbitltd.spoiwo.natures.xlsx

import com.norbitltd.spoiwo.model.Height._
import com.norbitltd.spoiwo.model.Width._
import com.norbitltd.spoiwo.model._
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions.{ convertSheet, writeToExistingSheet }
import com.norbitltd.spoiwo.natures.xlsx.Utils._
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel._
import org.scalatest.{ FlatSpec, Matchers }
import scala.collection.JavaConverters._
import scala.language.postfixOps

class Model2XlsxConversionsForSheetSpec extends FlatSpec with Matchers {

  private def workbook: XSSFWorkbook = new XSSFWorkbook()

  private def convert(s: Sheet): XSSFSheet = convertSheet(s, workbook)

  private def defaultSheet = convert(Sheet())

  "Sheet conversions" should "set sheet name to 'Sheet1' by default in empty workbook" in {
    defaultSheet.getSheetName shouldBe "Sheet1"
  }

  it should "set sheet name to 'Sheet 3'in a workbook with already 2 sheets defined" in {
    val w: XSSFWorkbook = workbook
    w.createSheet("Test1")
    w.createSheet("Test2")
    val sheet3: XSSFSheet = convertSheet(Sheet(), w)
    sheet3.getSheetName shouldBe "Sheet3"
  }

  it should "return defined name when set explicitly" in {
    val model = Sheet(name = "My Name")
    val xssf = convert(model)
    xssf.getSheetName shouldBe "My Name"
  }

  it should "write to existing sheet by overwriting existing cells" in {
    val w: XSSFWorkbook = workbook
    w.createSheet("Existing")
    val previousSheet = generateSheet(0 to 2, 0 to 5)
    val existingPoiSheet: XSSFSheet = convertSheet(previousSheet, w)
    val newSheet = generateSheet[Int, Int](1 to 3, 1 to 4, (rowNum, colNum) => s"NEW $colNum,$rowNum")
    writeToExistingSheet(newSheet, existingPoiSheet)
    val dataMatrix = mergeSheetData(Seq(previousSheet, newSheet), _.value.toString)
    nonEqualCells(existingPoiSheet, dataMatrix, _.getStringCellValue) shouldBe empty
  }

  it should "removes obsolete formulas" in {
    val inputStream = this.getClass.getResourceAsStream("/with_formula.xlsx")
    val w = XSSFWorkbookFactory.createWorkbook(inputStream)
    w.getCalculationChain.getCTCalcChain.sizeOfCArray() should equal(1)
    val existingPoiSheet = w.getSheetAt(0)
    val newSheet = Sheet(Row(Seq(Cell(2.0, index = 0)), index = 0))
    writeToExistingSheet(newSheet, existingPoiSheet)
    existingPoiSheet.getRow(0).getCell(0).getCellType should equal(CellType.NUMERIC)
    w.getCalculationChain.getCTCalcChain.sizeOfCArray() should equal(0)
  }

  it should "not have 1st column style by default" in {
    val xssfFont = defaultSheet.getColumnStyle(0).asInstanceOf[XSSFCellStyle].getFont
    xssfFont.getFontName shouldBe "Calibri"
    xssfFont.getFontHeightInPoints shouldBe 11
  }

  it should "have 1st column width when set with column without index" in {
    val model = Sheet(columns = Column(width = 200 unitsOfWidth) :: Nil)
    val xssf = convert(model)
    xssf.getColumnWidth(0) shouldBe 200
  }

  it should "have correct 1st and 2nd column style when set with column without index" in {
    val column1 = Column(width = 200 unitsOfWidth)
    val column2 = Column(width = 300 unitsOfWidth)
    val model = Sheet(columns = column1 :: column2 :: Nil)
    val xssf = convert(model)

    xssf.getColumnWidth(0) shouldBe 200
    xssf.getColumnWidth(1) shouldBe 300
  }

  it should "have 3rd column style when set with column with index" in {
    val model = Sheet(columns = Column(width = 300 unitsOfWidth, index = 2) :: Nil)
    val xssf = convert(model)
    xssf.getColumnWidth(2) shouldBe 300
  }

  it should "have correct 4th and 6th column style when set with column without index" in {
    val column1 = Column(width = 200 unitsOfWidth, index = 3)
    val column2 = Column(width = 300 unitsOfWidth, index = 5)
    val model = Sheet(columns = column1 :: column2 :: Nil)
    val xssf = convert(model)

    xssf.getColumnWidth(3) shouldBe 200
    xssf.getColumnWidth(4) shouldBe (8 characters).toUnits
    xssf.getColumnWidth(5) shouldBe 300
  }

  it should "not allow mixing columns with and without index" in {
    val column1 = Column(style = CellStyle(font = Font(fontName = "Arial", height = 14 points)))
    val column2 = Column(style = CellStyle(font = Font(fontName = "Arial", height = 16 points)), index = 5)
    val model = Sheet(columns = column1 :: column2 :: Nil)
    an[IllegalArgumentException] should be thrownBy convert(model)
  }

  it should "not allow 2 columns with the same index" in {
    val column1 = Column(style = CellStyle(font = Font(fontName = "Arial", height = 14 points)), index = 3)
    val column2 = Column(style = CellStyle(font = Font(fontName = "Arial", height = 16 points)), index = 5)
    val column3 = Column(style = CellStyle(font = Font(fontName = "Arial", height = 16 points)), index = 3)
    val model = Sheet(columns = column1 :: column2 :: column3 :: Nil)
    an[IllegalArgumentException] should be thrownBy convert(model)
  }

  it should "have no rows by default" in {
    defaultSheet.rowIterator().hasNext shouldBe false
  }

  it should "have 1st row with initialized with row and without index" in {
    val row = Row().withCellValues("Test", 34)
    val model = Sheet(rows = row :: Nil)

    val xlsx = convert(model)
    val xlsxRow = xlsx.getRow(0)
    xlsxRow.getCell(0).getStringCellValue shouldBe "Test"
    xlsxRow.getCell(1).getNumericCellValue shouldBe 34
  }

  it should "have 1st and 2nd row with initialized with row and without index" in {
    val row1 = Row().withCellValues("Test", 34)
    val row2 = Row().withCellValues("OK")
    val model = Sheet(rows = row1 :: row2 :: Nil)

    val xlsx = convert(model)
    val xlsxRow1 = xlsx.getRow(0)
    val xlsxRow2 = xlsx.getRow(1)
    xlsxRow1.getCell(0).getStringCellValue shouldBe "Test"
    xlsxRow1.getCell(1).getNumericCellValue shouldBe 34
    xlsxRow2.getCell(0).getStringCellValue shouldBe "OK"
  }

  it should "have 3rd row with initialized with row with index" in {
    val row = Row(index = 2).withCellValues("Test", 34)
    val model = Sheet(rows = row :: Nil)

    val xlsx = convert(model)
    val xlsxRow = xlsx.getRow(2)
    xlsxRow.getCell(0).getStringCellValue shouldBe "Test"
    xlsxRow.getCell(1).getNumericCellValue shouldBe 34
  }

  it should "have 3rd and 5th row with initialized with row and index" in {
    val row1 = Row(index = 4).withCellValues("Test", 34)
    val row2 = Row(index = 6).withCellValues("OK")
    val model = Sheet(rows = row1 :: row2 :: Nil)

    val xlsx = convert(model)
    val xlsxRow1 = xlsx.getRow(4)
    val xlsxRow2 = xlsx.getRow(6)
    xlsxRow1.getCell(0).getStringCellValue shouldBe "Test"
    xlsxRow1.getCell(1).getNumericCellValue shouldBe 34
    xlsxRow2.getCell(0).getStringCellValue shouldBe "OK"
  }

  it should "have image in the correct position" in {
    val row1 = Row(index = 0)
    val row2 = Row(index = 1)
    val row3 = Row(index = 2)
    val model = Sheet(rows = row1 :: row2 :: row3 :: Nil, images = Image(CellRange(2 -> 2, 2 -> 2), "src/test/resources/einstein.jpg") :: Nil)
      .withColumns(Column(index = 0), Column(index = 1))

    val xlsx = convert(model)

    val drawing = xlsx.createDrawingPatriarch()
    val pictures = drawing.getShapes.asScala.toStream.collect {
      case p: XSSFPicture => p
    }

    pictures.size shouldBe 1

    pictures.head.getClientAnchor.getRow1 shouldBe 2
    pictures.head.getClientAnchor.getCol1 shouldBe 2
  }

  it should "not allow duplicate row indexes" in {
    val row1 = Row(index = 4).withCellValues("Test", 34)
    val row2 = Row(index = 6).withCellValues("OK")
    val row3 = Row(index = 4).withCellValues("Override")
    val model = Sheet(rows = row1 :: row2 :: row3 :: Nil)
    an[IllegalArgumentException] should be thrownBy convert(model)
  }

  it should "not allow mixing indexed and non indexed rows" in {
    val row1 = Row(index = 4).withCellValues("Test", 34)
    val row2 = Row().withCellValues("OK")
    val model = Sheet(rows = row1 :: row2 :: Nil)
    an[IllegalArgumentException] should be thrownBy convert(model)
  }

}
