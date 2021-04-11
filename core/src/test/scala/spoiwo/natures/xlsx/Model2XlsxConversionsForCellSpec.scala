package spoiwo.natures.xlsx

import java.util.{Calendar, Date}
import java.time.{LocalDate, LocalDateTime, ZoneId}
import Model2XlsxConversions.{convertCell, _}
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFWorkbook}

import scala.language.postfixOps
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers
import spoiwo.model.Height.HeightEnrichment
import spoiwo.model._

class Model2XlsxConversionsForCellSpec extends AnyFlatSpec with Matchers {

  private val defaultCell: XSSFCell = convert(Cell.Empty)

  private def convert(cell: Cell): XSSFCell = convertCell(Sheet(), Map(), Row(), cell, row)

  private def row = new XSSFWorkbook().createSheet().createRow(0)

  "Cell conversion" should "return blank cell type with empty string by default" in {
    defaultCell.getCellType shouldBe CellType.BLANK
    defaultCell.getStringCellValue shouldBe ""
  }

  it should "return a 0 column index by default if no other cells specified in the row" in {
    defaultCell.getColumnIndex shouldBe 0
  }

  it should "return cell style with 11pt Calibri by default" in {
    defaultCell.getCellStyle.getFont.getFontHeightInPoints shouldBe 11
    defaultCell.getCellStyle.getFont.getFontName shouldBe "Calibri"
  }

  it should "return cell style with 14pt Arial when explicitly specified" in {
    val cellStyle = CellStyle(font = Font(fontName = "Arial", height = 14 points))
    val model = Cell.Empty.withStyle(cellStyle)
    val xlsx = convert(model)

    xlsx.getCellStyle.getFont.getFontHeightInPoints shouldBe 14
    xlsx.getCellStyle.getFont.getFontName shouldBe "Arial"
  }

  it should "return index of 3 when explicitly specified" in {
    val model = Cell.Empty.withIndex(3)
    val xlsx = convert(model)

    xlsx.getColumnIndex shouldBe 3
  }

  it should "return index of 2 when row has already 2 other cells" in {
    val row = new XSSFWorkbook().createSheet().createRow(0)
    row.createCell(0)
    row.createCell(1)

    val model = Cell.Empty
    val xlsx = convertCell(Sheet(), Map(), Row(), model, row)
    xlsx.getColumnIndex shouldBe 2
  }

  it should "return string cell when set up with 'String'" in {
    val model = Cell("TEST_STRING")
    val xlsx = convert(model)
    xlsx.getCellType shouldBe CellType.STRING
    xlsx.getStringCellValue shouldBe "TEST_STRING"
  }

  it should "return string cell when set up with String with newline value" in {
    val model = Cell("TEST_STRING\nAnd a 2nd line")
    val xlsx = convert(model)
    xlsx.getCellType shouldBe CellType.STRING
    xlsx.getStringCellValue shouldBe "TEST_STRING\nAnd a 2nd line"
  }

  it should "return formula cell when set up with string starting with '=' sign" in {
    val model = Cell("=1000/3+7")
    val xlsx = convert(model)
    xlsx.getCellType shouldBe CellType.FORMULA
    xlsx.getCellFormula shouldBe "1000/3+7"
  }

  it should "return numeric cell when set up with double value" in {
    val model = Cell(90.45)
    val xlsx = convert(model)
    xlsx.getCellType shouldBe CellType.NUMERIC
    xlsx.getNumericCellValue shouldBe 90.45
  }

  it should "return numeric cell when set up with big decimal value" in {
    val model = Cell(BigDecimal(90.45))
    val xlsx = convert(model)
    xlsx.getCellType shouldBe CellType.NUMERIC
    xlsx.getNumericCellValue shouldBe 90.45
  }

  it should "return numeric cell when set up with int value" in {
    val model = Cell(90)
    val xlsx = convert(model)
    xlsx.getCellType shouldBe CellType.NUMERIC
    xlsx.getNumericCellValue shouldBe 90
  }

  it should "return numeric cell when set up with long value" in {
    val model = Cell(10000000000000L)
    val xlsx = convert(model)
    xlsx.getCellType shouldBe CellType.NUMERIC
    xlsx.getNumericCellValue shouldBe 10000000000000L
  }

  it should "return boolean cell when set up with boolean value" in {
    val model = Cell(true)
    val xlsx = convert(model)
    xlsx.getCellType shouldBe CellType.BOOLEAN
    xlsx.getBooleanCellValue shouldBe true
  }

  it should "return numeric cell when set up with java.util.Date value" in {
    val localDate = LocalDate.of(2011, 11, 13)
    val model = Cell(Date.from(localDate.atStartOfDay().atZone(ZoneId.systemDefault()).toInstant))
    val xlsx = convert(model)

    val date = xlsx.getDateCellValue.toInstant.atZone(ZoneId.systemDefault()).toLocalDate
    date.getYear shouldBe 2011
    date.getMonthValue shouldBe 11
    date.getDayOfMonth shouldBe 13
  }

  it should "return numeric cell when set up with java.util.Calendar value" in {
    val calendar = Calendar.getInstance()
    calendar.set(2011, 11, 13)
    val model = Cell(calendar)
    val xlsx = convert(model)

    val date = xlsx.getDateCellValue.toInstant.atZone(ZoneId.systemDefault()).toLocalDate
    date.getYear shouldBe 2011
    date.getMonthValue shouldBe 12
    date.getDayOfMonth shouldBe 13
  }

  it should "return numeric cell when set up with java.time.LocalDate value" in {
    test(LocalDate.of(2011, 6, 13))
    test(LocalDate.of(2011, 11, 13))

    def test(ld: LocalDate): Unit = {
      val model = Cell(ld)
      val xlsx = convert(model)

      val date = xlsx.getDateCellValue.toInstant.atZone(ZoneId.systemDefault())
      date.getYear shouldBe ld.getYear
      date.getMonthValue shouldBe ld.getMonthValue
      date.getDayOfMonth shouldBe ld.getDayOfMonth
      date.getHour shouldBe 0
      date.getMinute shouldBe 0
      date.getSecond shouldBe 0
    }
  }

  it should "return numeric cell when set up with java.time.LocalDateTime value" in {
    test(LocalDateTime.of(2011, 6, 13, 15, 30, 10))
    test(LocalDateTime.of(2011, 11, 13, 15, 30, 10))

    def test(ldt: LocalDateTime): Unit = {
      val model = Cell(ldt)
      val xlsx = convert(model)

      val date = xlsx.getDateCellValue.toInstant.atZone(ZoneId.systemDefault())
      date.getYear shouldBe ldt.getYear
      date.getMonthValue shouldBe ldt.getMonthValue
      date.getDayOfMonth shouldBe ldt.getDayOfMonth
      date.getHour shouldBe ldt.getHour
      date.getMinute shouldBe ldt.getMinute
      date.getSecond shouldBe ldt.getSecond
    }
  }

  it should "return string cell with the date formatted yyyy-MM-dd if date before 1904" in {
    val model = Cell(LocalDate.of(1856, 11, 3))
    val xlsx = convert(model)
    "1856-11-03" shouldBe xlsx.getStringCellValue
  }

  it should "apply 14pt Arial cell style for column when set explicitly" in {
    val column = Column(index = 0, style = CellStyle(font = Font(fontName = "Arial", height = 14 points)))
    val sheet = Sheet(Row(Cell("Test"))).withColumns(column)
    val xlsx = sheet.convertAsXlsx()

    val cellStyle = xlsx.getSheetAt(0).getRow(0).getCell(0).getCellStyle
    cellStyle.getFont.getFontName shouldBe "Arial"
    cellStyle.getFont.getFontHeightInPoints shouldBe 14
  }

  it should "return a string cell with hyperlink when setup with HyperLinkUrl value" in {
    val model = Cell(HyperLink("View Item", "https://www.google.com"))
    val xlsx = convert(model)
    xlsx.getCellType shouldBe CellType.STRING
    xlsx.getHyperlink.getAddress shouldBe "https://www.google.com"
  }
}
