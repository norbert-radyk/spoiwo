package spoiwo.examples.gettingstarted

import spoiwo.model.{Row, Sheet}
import spoiwo.natures.xlsx.Model2XlsxConversions.XlsxSheet

object HelloWorld {

  def main(args: Array[String]): Unit = {
    val helloWorldSheet = Sheet(name = "Hello Sheet", row = Row().withCellValues("Hello World!"))

    helloWorldSheet.saveAsXlsx(args(0))
  }
}
