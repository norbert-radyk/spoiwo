package com.norbitltd.spoiwo.examples.gettingstarted

import com.norbitltd.spoiwo.model._
import org.joda.time.LocalDate
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._
import com.norbitltd.spoiwo.model.enums.CellFill

object GettingStartedExample {

  val headerStyle =
    CellStyle(fillPattern = CellFill.Solid, fillForegroundColor = Color.AquaMarine, fillBackgroundColor = Color.AquaMarine, font = Font(bold = true))

  val rowHeight = new Height(1000, HeightUnit.Unit)

  val einsteinImage = Image(CellRange(2 -> 2, 4 -> 5), "examples/src/main/resources/einstein.jpg")

  val gettingStartedSheet: Sheet = Sheet(name = "Some serious stuff", images = List(einsteinImage))
    .withRows(
      Row(style = headerStyle).withCellValues("NAME", "BIRTH DATE", "DIED AGED", "FEMALE"),
      Row().withCellValues("Marie Curie", new LocalDate(1867, 11, 7), 66, true),
      Row().withCellValues("Albert Einstein", new LocalDate(1879, 3, 14), 76, false),
      Row().withCellValues("Erwin Shrodinger", new LocalDate(1887, 8, 12), 73, false)
    )
    .withColumns(
      Column(index = 0, style = CellStyle(font = Font(bold = true)), autoSized = true)
    )

  def main(args: Array[String]) {
    gettingStartedSheet.saveAsXlsx(args(0))
  }
}
