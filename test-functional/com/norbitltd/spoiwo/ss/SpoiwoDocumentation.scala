package com.norbitltd.spoiwo.ss

import org.apache.poi.hssf.usermodel.HeaderFooter

class SpoiwoDocumentation {

  def usingNewLinesInCells() {
    val row = Row(
      heightInPoints = 20,
      cells = Vector(Cell.Empty, Cell("Use \n with word wrap on to create a new line"))
    )

    Sheet(Row.Empty, row).saveAsXlsx("ooxml-newlines.xlsx")
  }

  def dataFormats() {
    def createFormattedCellRow(value: Double, format: String) =
      Row(Cell(value, CellStyle(dataFormat = CellDataFormat(format))))

    Sheet(
      name = "format sheet",
      rows = createFormattedCellRow(11111.25, "0.0") ::
        createFormattedCellRow(11111.25, "#,##0.0000") :: Nil
    ).saveAsXlsx("workbook.xls")
  }

  def fitSheetToOnePage() {

    val printSetup = PrintSetup.Default
      .withFitHeight(1)
      .withFitWidth(1)

    Sheet(name = "format sheet", printSetup = printSetup)
      .saveAsXlsx("workbook.xlsx")
  }

  def setPageNumbersOnFooter() {
    Sheet(name = "format sheet", footer = UnifiedFooter(pageFooterData =
      FooterData(right = "Page " + HeaderFooter.page + " of " + HeaderFooter.numPages))
    ).saveAsXlsx("workbook.xlsx")
  }

}