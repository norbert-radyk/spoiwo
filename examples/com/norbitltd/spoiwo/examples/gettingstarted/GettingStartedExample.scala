package com.norbitltd.spoiwo.examples.gettingstarted

import com.norbitltd.spoiwo.model._
import org.joda.time.{LocalDate}
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._

object GettingStartedExample {

  val sheet1 = Sheet(name = "Hello Sheet",
    row = Row.Empty.withCellValues("Hello World!")
  )

  val headerStyle =
    CellStyle(fillPattern = CellFill.Solid, fillForegroundColor = Color.AquaMarine, fillBackgroundColor = Color.AquaMarine, font = Font(bold = true))

  val sheet2 = Sheet(name = "Some serious stuff")
    .withColumns(
      Column(index = 1, style = CellStyle(dataFormat = CellDataFormat("dd MMM yyyy")))
    )
    .withRows(
      Row.Empty.withCellValues("NAME", "BIRTH DATE", "DIED AGED", "FEMALE").withStyle(headerStyle),
      Row.Empty.withCellValues("Marie Curie", new LocalDate(1867, 11, 7), 66, true),
      Row.Empty.withCellValues("Albert Einstein", new LocalDate(1879, 3, 14), 76, false),
      Row.Empty.withCellValues("Erwin Shrodinger", new LocalDate(1887, 8, 12), 73, false)
    )

  def main(args: Array[String]) {
    sheet1.saveAsXlsx("C:\\Reports\\hello_world.xlsx")
    sheet2.saveAsXlsx("C:\\Reports\\getting_started.xlsx")
  }

}
