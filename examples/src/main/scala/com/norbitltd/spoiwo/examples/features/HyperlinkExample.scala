package com.norbitltd.spoiwo.examples.features

import com.norbitltd.spoiwo.model.enums.HyperLinkType
import com.norbitltd.spoiwo.model.{HyperLink, Row, Sheet}
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._

object HyperlinkExample {

  val hyperlinksSheet: Sheet = Sheet(name = "Hyperlinks")
    .withRows(
      Row().withCellValues("PERSON", "EMAIL"),
      Row().withCellValues(
        HyperLink(text = "Bill Gates", address = "https://pl.wikipedia.org/wiki/Bill_Gates"),
        HyperLink(text = "bill.gates@microsoft.com", address = "mailto:bill.gates@microsoft.com", linkType = HyperLinkType.Email)
      )
    )

  def main(args: Array[String]): Unit = {
    hyperlinksSheet.saveAsXlsx(args(0))
  }

}
