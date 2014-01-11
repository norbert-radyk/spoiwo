package com.norbitltd.spoiwo.excel

class T008Test {

  def mergingCells() {
    Sheet(
      name = "new sheet",
      rows = Row(Cell("This is a test of merging")) :: Nil,
      mergedRegions = CellRangeAddress(rowRange = 1 -> 1, columnRange = 1 -> 2) :: Nil
    ).save("workbook.xlsx")
  }
}
