package com.norbitltd.spoiwo.examples.gettingstarted

import com.norbitltd.spoiwo.model._
import org.joda.time.LocalDate
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._
import com.norbitltd.spoiwo.model.enums.CellFill

object GettingStartedExample {

  val sheet1 = Sheet(name = "Hello Sheet",
    row = Row.Empty.withCellValues("Hello World!")
  )

  val headerStyle =
    CellStyle(fillPattern = CellFill.Solid, fillForegroundColor = Color.AquaMarine, fillBackgroundColor = Color.AquaMarine, font = Font(bold = true))

  val sheet2 = Sheet(name = "Some serious stuff")
    .withRows(
      Row(style = headerStyle).withCellValues(headerStyle, "NAME", "BIRTH DATE", "DIED AGED", "FEMALE"),
      Row.Empty.withCellValues("Marie Curie", new LocalDate(1867, 11, 7), 66, true),
      Row.Empty.withCellValues("Albert Einstein", new LocalDate(1879, 3, 14), 76, false),
      Row.Empty.withCellValues("Erwin Shrodinger", new LocalDate(1887, 8, 12), 73, false)
    )
    .withColumns(
      Column(index = 0, style = CellStyle(font = Font(bold = true)), autoSized = true)
    )

  def main(args: Array[String]) {
    sheet1.saveAsXlsx("C:\\Reports\\hello_world.xlsx")
    sheet2.saveAsXlsx("C:\\Reports\\getting_started.xlsx")
  }

}
