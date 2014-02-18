package com.norbitltd.spoiwo.model

class T008Test {

  def mergingCells() {
    Sheet(
      name = "new sheet",
      rows = Row(Cell("This is a test of merging")) :: Nil,
      mergedRegions = CellRange(rowRange = 1 -> 1, columnRange = 1 -> 2) :: Nil
    ).saveAsXlsx("workbook.xlsx")
  }
}
