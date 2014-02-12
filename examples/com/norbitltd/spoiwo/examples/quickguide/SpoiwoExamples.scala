package com.norbitltd.spoiwo.examples.quickguide

import com.norbitltd.spoiwo.ss._
import org.apache.poi.ss.util.WorkbookUtil
import java.util.{Calendar, Date}
import org.apache.poi.ss.usermodel.{HorizontalAlignment => HA, VerticalAlignment => VA, FillPatternType, BorderStyle}

class SpoiwoExamples {

  def newWorkbook() = Workbook().saveAsXlsx("workbook.xlsx")

  def newSheet() = Workbook(
    Sheet(name = "new sheet"),
    Sheet(name = "second sheet"),
    Sheet(name = WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]"))
  ).saveAsXlsx("workbook.xlsx")

  def creatingCells() = Sheet(name = "new sheet",
    row = Row().withCellValues(1, 1.2, "This is a string", true)
  ).saveAsXlsx("workbook.xlsx")

  def creatingDateCells() {
    val style = CellStyle(dataFormat = CellDataFormat("m/d/yy h:mm"))
    val simplestDateCell = Cell(new Date())
    val formattedDateCell = Cell(new Date(), style)
    val formattedCalendarCell = Cell(Calendar.getInstance(), style)

    Sheet(name = "new sheet",
      row = Row(simplestDateCell, formattedDateCell, formattedCalendarCell)
    ).saveAsXlsx("workbook.xlsx")
  }

  //There is no way to explicitly specify type of the cell as it is derived form the type of value passed
  //and therefore the error cell type can't be specified in the way shown in Java POI example.
  //I hesitate to believe it makes much sense to explicitly generate error cells when generating the report,
  //however to simulate this case we're passing error formula here.
  def workingWithDifferentTypesOfCells() = Sheet(name = "new sheet",
    row = Row(index = 2).withCellValues(1.1, new Date(), Calendar.getInstance(), "a string", true, "=1/0")
  ).saveAsXlsx("workbook.xlsx")


  def createCell(ha: HA, va: VA) =
    Cell("Align It", style = CellStyle(horizontalAlignment = ha, verticalAlignment = va))

  def variousAlignmentOptions() {
    val alignments = List(HA.CENTER -> VA.BOTTOM, HA.CENTER_SELECTION -> VA.BOTTOM, HA.FILL -> VA.CENTER,
      HA.GENERAL -> VA.CENTER, HA.JUSTIFY -> VA.JUSTIFY, HA.LEFT -> VA.TOP, HA.RIGHT -> VA.TOP)

    Sheet(
      Row(index = 2, heightInPoints = 30).withCells(alignments.map((createCell _).tupled))
    ).saveAsXlsx("xssf-align.xlsx")
  }

  def workingWithBorders() {
    val borders = CellBorders(
      bottomStyle = BorderStyle.THIN, bottomColor = Color.BLACK,
      leftStyle = BorderStyle.THIN, leftColor = Color.GREEN,
      rightStyle = BorderStyle.THIN, rightColor = Color.BLUE,
      topStyle = BorderStyle.MEDIUM_DASHED, topColor = Color.BLACK
    )

    Sheet(name = "new sheet",
      row = Row(index = 1, Cell(value = 4, index = 1, style = CellStyle(borders = borders)))
    ).saveAsXlsx("workbook.xls")
  }

  def fillsAndColors() = Sheet(name = "new sheet",
    row = Row(index = 1,
      Cell.Empty,
      Cell("X", CellStyle(fillBackgroundColor = Color.AQUA, fillPattern = FillPatternType.BIG_SPOTS)),
      Cell("X", CellStyle(fillForegroundColor = Color.ORANGE, fillPattern = FillPatternType.SOLID_FOREGROUND))
    )
  ).saveAsXlsx("workbook.xls")

  def mergingCells() = Sheet(name = "new sheet")
    .withRows(Row(index = 1, Cell("This is a test of merging", index = 1)))
    .withMergedRegions(CellRange(1 -> 1, 1 -> 2))
    .saveAsXlsx("workbook.xls")

  def workingWithFonts() {
    val style = CellStyle(font = Font(heightInPoints = 24, fontName = "Courier New", italic = true, strikeout = true))
    Sheet(name = "new sheet",
      row = Row(index = 1, Cell("This is a test of fonts", 1, style))
    ).saveAsXlsx("workbook.xls")
  }

  def workingWithMultipleFonts() {
    val cell = Cell(value = 0, style = CellStyle(font = Font(bold = true)))
    Sheet(name = "new sheet",
      rows = (0 to 10000).map(i => Row(cell)).toList
    ).saveAsXlsx("workbook.xls")
  }

  def customColors() = Sheet(Row(Cell("custom XSSF colors", style =
    CellStyle(fillForegroundColor = Color(128, 0, 128), fillPattern = FillPatternType.SOLID_FOREGROUND)
  )))

  def usingNewLinesInCells() = {
    val cell = Cell("Use \n with word wrap on to create a new line", index = 2)
    val row = Row(index = 2, heightInPoints = 2 * Row.DefaultHeightInPoints, cells = List(cell))
    val sheet = Sheet(row).withColumns(Column(2, autoSized = true))
    sheet.saveAsXlsx("ooxml-newlines.xlsx")
  }

  def dataFormats() = Sheet(name = "format sheet",
    Row(Cell(11111.25, CellStyle(dataFormat = CellDataFormat("0.0")))),
    Row(Cell(11111.25, CellStyle(dataFormat = CellDataFormat("#,##0.0000"))))
  ).saveAsXlsx("workbook.xls")

  //TODO Refactor print setup
  def fitSheetToOnePage() = Sheet(name = "format sheet",
    properties = SheetProperties(autoBreaks = true),
    printSetup = PrintSetup(fitWidth = Some(1.toShort), fitHeight = Some(1.toShort))
  ).saveAsXlsx("workbook.xls")

}
