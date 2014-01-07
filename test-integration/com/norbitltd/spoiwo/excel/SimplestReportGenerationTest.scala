package com.norbitltd.spoiwo.excel

object SimplestReportGenerationTest extends AbstractReportGenerator {

  override def getWorkbook = Workbook(
    Sheet(
      name = "Test Sheet",
      rows = List(
        Row(
          Cell("Test string value", CellStyle.Default)
        ))
    )
  )

}
