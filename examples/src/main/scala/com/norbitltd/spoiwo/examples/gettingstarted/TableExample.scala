package com.norbitltd.spoiwo.examples.gettingstarted

import com.norbitltd.spoiwo.model._
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._

object TableExample {

  def main(args: Array[String]): Unit =  {

    val sheetWithTable = Sheet(name = "Table Sheet",
      rows = {
        val headerRow = Row().withCellValues("a", "b", "c")
        val rows = (1 to 10).map(i => Row().withCellValues(i*1, i*2, i*3))
        headerRow :: rows.toList
      },
      tables = List(
        Table(
          cellRange = CellRange(0 -> 10, 0-> 2),
          style = TableStyle(TableStyleName.TableStyleMedium2, showRowStripes = true),
          enableAutoFilter = true
        )
      )
    )

    sheetWithTable.saveAsXlsx(args(0))
  }
}
