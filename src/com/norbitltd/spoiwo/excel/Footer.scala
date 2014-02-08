package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFSheet

object FooterData {
  val Empty = FooterData()
}

case class FooterData(left: String = "",
                      center: String = "",
                      right: String = "")

object Footer {

  val None = UnifiedFooter()

}

sealed trait Footer {

  def applyTo(sheet: XSSFSheet)

  private[excel] def applyData(footer: org.apache.poi.ss.usermodel.Footer, data: FooterData) {
    if (data.left != "") footer.setLeft(data.left)
    if (data.center != "") footer.setCenter(data.center)
    if (data.right != "") footer.setRight(data.right)
  }
}

case class UnifiedFooter(pageFooterData: FooterData = FooterData.Empty,
                         firstPageFooterData: FooterData = FooterData.Empty) extends Footer {

  def applyTo(sheet: XSSFSheet) {
    if (pageFooterData != FooterData.Empty) applyData(sheet.getFooter, pageFooterData)
    if (firstPageFooterData != FooterData.Empty) applyData(sheet.getFirstFooter, firstPageFooterData)
  }
}

case class OddEvenPageFooter(oddPageFooterData: FooterData = FooterData.Empty,
                             evenPageFooterData: FooterData = FooterData.Empty,
                             firstPageFooterData: FooterData = FooterData.Empty) extends Footer {

  def applyTo(sheet: XSSFSheet) {
    if (oddPageFooterData != FooterData.Empty) applyData(sheet.getOddFooter, oddPageFooterData)
    if (evenPageFooterData != FooterData.Empty) applyData(sheet.getEvenFooter, evenPageFooterData)
    if (firstPageFooterData != FooterData.Empty) applyData(sheet.getFirstFooter, firstPageFooterData)
  }
}
