package com.norbitltd.spoiwo.examples.gettingstarted

import com.norbitltd.spoiwo.model.{Row, Sheet}
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._

object HelloWorld {

  def main(args: Array[String]): Unit =  {
    val helloWorldSheet = Sheet(name = "Hello Sheet",
      row = Row().withCellValues("Hello World!")
    )

    helloWorldSheet.saveAsXlsx(args(0))
  }
}
