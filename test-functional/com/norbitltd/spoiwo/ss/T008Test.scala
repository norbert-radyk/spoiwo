package com.norbitltd.spoiwo.ss

class T008Test {

  def mergingCells() {
    Sheet(
      name = "new sheet",
      rows = Row(Cell("This is a test of merging")) :: Nil,
      mergedRegions = CellRange(rowRange = 1 -> 1, columnRange = 1 -> 2) :: Nil
    ).saveXLSX("workbook.xlsx")
  }
}
