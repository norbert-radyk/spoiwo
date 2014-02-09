package com.norbitltd.spoiwo.ss

object SimplestReportGenerationTest extends AbstractReportGenerator {

  override def getWorkbook = Workbook(
    Sheet("Test Sheet",
      Row().withCellValues("Test string value", 9, "000"),
      Row().withCellValues("Test string value2")
    )
  )

}
