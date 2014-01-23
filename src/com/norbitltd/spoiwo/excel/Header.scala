package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFSheet

object HeaderData {
  val Empty = HeaderData()
}

case class HeaderData(left: String = "",
                      center: String = "",
                      right: String = "")

object Header {

  val None = UnifiedHeader()

}

sealed trait Header {

  def apply(sheet: XSSFSheet)

  private[excel] def applyData(Header: org.apache.poi.ss.usermodel.Header, data: HeaderData) {
    if (data.left != "") Header.setLeft(data.left)
    if (data.center != "") Header.setCenter(data.center)
    if (data.right != "") Header.setRight(data.right)
  }
}

case class UnifiedHeader(pageHeaderData: HeaderData = HeaderData.Empty,
                         firstPageHeaderData: HeaderData = HeaderData.Empty) extends Header {

  def apply(sheet: XSSFSheet) {
    if (pageHeaderData != HeaderData.Empty) applyData(sheet.getHeader, pageHeaderData)
    if (firstPageHeaderData != HeaderData.Empty) applyData(sheet.getFirstHeader, firstPageHeaderData)
  }
}

case class OddEvenPageHeader(oddPageHeaderData: HeaderData = HeaderData.Empty,
                             evenPageHeaderData: HeaderData = HeaderData.Empty,
                             firstPageHeaderData: HeaderData = HeaderData.Empty) extends Header {

  def apply(sheet: XSSFSheet) {
    if (oddPageHeaderData != HeaderData.Empty) applyData(sheet.getOddHeader, oddPageHeaderData)
    if (evenPageHeaderData != HeaderData.Empty) applyData(sheet.getEvenHeader, evenPageHeaderData)
    if (firstPageHeaderData != HeaderData.Empty) applyData(sheet.getFirstHeader, firstPageHeaderData)
  }
}
