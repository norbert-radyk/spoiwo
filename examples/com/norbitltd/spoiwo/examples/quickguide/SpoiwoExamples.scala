package com.norbitltd.spoiwo.examples.quickguide

import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._
import com.norbitltd.spoiwo.model._
import org.apache.poi.ss.util.WorkbookUtil
import java.util.{Calendar, Date}

import com.norbitltd.spoiwo.model.{CellHorizontalAlignment => HA}
import com.norbitltd.spoiwo.model.{CellVerticalAlignment => VA}

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
    val alignments = List(HA.Center -> VA.Bottom, HA.CenterSelection -> VA.Bottom, HA.Fill -> VA.Center,
      HA.General -> VA.Center, HA.Justify -> VA.Justify, HA.Left -> VA.Top, HA.Right -> VA.Top)

    Sheet(
      Row(index = 2, heightInPoints = 30).withCells(alignments.map((createCell _).tupled))
    ).saveAsXlsx("xssf-align.xlsx")
  }


  def workingWithBorders() {
    val borders = CellBorders(
      bottomStyle = CellBorderStyle.Thin, bottomColor = Color.Black,
      leftStyle = CellBorderStyle.Thin, leftColor = Color.Green,
      rightStyle = CellBorderStyle.Thin, rightColor = Color.Blue,
      topStyle = CellBorderStyle.MediumDashed, topColor = Color.Black
    )

    Sheet(name = "new sheet",
      row = Row(index = 1, Cell(value = 4, index = 1, style = CellStyle(borders = borders)))
    ).saveAsXlsx("workbook.xls")
  }


  def fillsAndColors() = Sheet(name = "new sheet",
    row = Row(index = 1,
      Cell.Empty,
      Cell("X", CellStyle(fillBackgroundColor = Color.Aqua, fillPattern = CellFill.Pattern.BigSpots)),
      Cell("X", CellStyle(fillForegroundColor = Color.Orange, fillPattern = CellFill.Solid))
    )
  ).saveAsXlsx("workbook.xls")


  def mergingCells() = Sheet(name = "new sheet")
    .withRows(Row(index = 1, Cell("This is a test of merging", index = 1)))
    .withMergedRegions(CellRange(1 -> 1, 1 -> 2))
    .saveAsXlsx("workbook.xls")


  def workingWithFonts() {
    import Measure._
    val style = CellStyle(font = Font(height = 24 points, fontName = "Courier New", italic = true, strikeout = true))
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
    CellStyle(fillForegroundColor = Color(128, 0, 128), fillPattern = CellFill.Solid)
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


  def fitSheetToOnePage() = Sheet(name = "format sheet",
    properties = SheetProperties(autoBreaks = true),
    printSetup = PrintSetup(fitWidth = 1, fitHeight = 1)
  ).saveAsXlsx("workbook.xls")


  def setPrintArea() = Sheet(name = "Sheet1",
    properties = SheetProperties(printArea = CellRange(0 -> 0, 0 -> 1))
  ).saveAsXlsx("workbook.xlsx")


  def setPageNumbersOnFooter() = Sheet(name = "format sheet",
    footer = Footer.Standard(right = "Page &P of &N")
  ).saveAsXlsx("workbook.xlsx")


  def setSheetAsSelected() = Sheet(name = "row sheet", properties = SheetProperties(selected = true))


  def setZoomMagnification() = Sheet(name = "new sheet", properties = SheetProperties(zoom = 75))


  def splitAndFreezePanes() = Workbook(
      //INFO: When creating freeze pane scrolling default to freeze, so no need duplicating 0, 1
      Sheet(name = "new sheet", paneAction = FreezePane(0, 1)),
      Sheet(name = "second sheet", paneAction = FreezePane(1, 0)),
      Sheet(name = "third sheet", paneAction = FreezePane(2, 2)),
      Sheet(name = "fourth sheet", paneAction = SplitPane(2000, 2000, 0, 0, Pane.LowerLeftPane))
  ).saveAsXlsx("workbook.xlsx")

  def repeatingRowsAndColumns() = Workbook(
    Sheet(name = "Sheet1", repeatingRows = RowRange(4 -> 5)),
    Sheet(name = "Sheet2", repeatingColumns = ColumnRange("A" -> "C"))
  ).saveAsXlsx("workbook.xlsx")

  def headersAndFooters() = Sheet(name = "new sheet",
    header = Header.Standard(left = "Left Header", center = "Center Header",
      right = """&"Stencil-Normal,Italic"&16 Right w / Stencil - Normal Italic font and size 16""")
  ).saveAsXlsx("workbook.xlsx")
}