package com.norbitltd.spoiwo.html

import SS2HtmlConversions._
import com.norbitltd.spoiwo.ss.{Sheet, Row}

class SS2HTMLConversionsTest {

  def testImplicits() = {
    val sheet = Sheet(
      Row().withCellValues(34, 123, "Italy"),
      Row().withCellValues(43, 22.9, "Spain")
    )

   sheet.asHtmlTableElem()
  }

}
