package com.norbitltd.spoiwo.natures.html

import Model2HtmlConversions._
import com.norbitltd.spoiwo.model.{Sheet, Row}

class SS2HTMLConversionsTest {

  def testImplicits() = {
    val sheet = Sheet(
      Row().withCellValues(34, 123, "Italy"),
      Row().withCellValues(43, 22.9, "Spain")
    )

   sheet.asHtmlTableElem()
  }

}
