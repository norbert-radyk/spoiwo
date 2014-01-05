package com.norbitltd.spoiwo.excel

object SimplestReportGenerationTest extends AbstractReportGenerator {

  override def getWorkbook = XWorkbook(
    XSheet(
      name = "Test Sheet",
      rows = List(
        XRow(
          XCell("Test string value", XCellStyle.Default)
        ))
    )
  )

}
