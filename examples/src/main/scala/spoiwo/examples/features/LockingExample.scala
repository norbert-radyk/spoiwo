package spoiwo.examples.features

import spoiwo.model._
import spoiwo.natures.xlsx.Model2XlsxConversions.XlsxSheet

import java.time.LocalDate

object LockingExample {

  val locking: SheetLocking = SheetLocking(
    lockedInsertColumns = true,
    lockedInsertRows = true,
    lockedDeleteColumns = true,
    lockedDeleteRows = true
  )

  val gettingStartedSheet: Sheet = Sheet(name = "Locked insert and delete")
    .withRows(
      Row().withCellValues("NAME", "BIRTH DATE"),
      Row().withCellValues("Marie Curie", LocalDate.of(1867, 11, 7)),
      Row().withCellValues("Albert Einstein", LocalDate.of(1879, 3, 14)),
      Row().withCellValues("Erwin Shrodinger", LocalDate.of(1887, 8, 12))
    )
    .withColumns(
      Column(index = 0, style = CellStyle(font = Font(bold = true)), autoSized = true)
    )
    .withLocking(locking)

  def main(args: Array[String]): Unit = {
    gettingStartedSheet.saveAsXlsx(args(0))
  }

}
