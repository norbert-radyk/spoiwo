package spoiwo.examples.features

import spoiwo.model._
import spoiwo.model.enums.HyperLinkType
import spoiwo.natures.streaming.xlsx.Model2XlsxConversions.XlsxSheet

object HyperlinkExample {

  val hyperlinksSheet: Sheet = Sheet(name = "Hyperlinks")
    .withRows(
      Row().withCellValues("PERSON", "EMAIL"),
      Row().withCellValues(
        HyperLink(text = "Bill Gates", address = "https://pl.wikipedia.org/wiki/Bill_Gates"),
        HyperLink(
          text = "bill.gates@microsoft.com",
          address = "mailto:bill.gates@microsoft.com",
          linkType = HyperLinkType.Email
        )
      )
    )

  def main(args: Array[String]): Unit = {
    hyperlinksSheet.saveAsXlsx(args(0))
  }

}
