package com.norbitltd.spoiwo.model

import org.apache.poi.ss.usermodel.Sheet._
import org.apache.poi.xssf.usermodel.XSSFSheet

object Margins extends Factory {

  private lazy val defaultLeft = defaultPOISheet.getMargin(LeftMargin)
  private lazy val defaultRight = defaultPOISheet.getMargin(RightMargin)
  private lazy val defaultTop = defaultPOISheet.getMargin(TopMargin)
  private lazy val defaultBottom = defaultPOISheet.getMargin(BottomMargin)
  private lazy val defaultHeader = defaultPOISheet.getMargin(HeaderMargin)
  private lazy val defaultFooter = defaultPOISheet.getMargin(FooterMargin)

  val Default = Margins()

  def apply(top: Double = defaultTop,
            bottom: Double = defaultBottom,
            right: Double = defaultRight,
            left: Double = defaultLeft,
            header: Double = defaultHeader,
            footer: Double = defaultFooter): Margins =
    Margins(
      wrap(top, defaultTop),
      wrap(bottom, defaultBottom),
      wrap(right, defaultRight),
      wrap(left, defaultLeft),
      wrap(header, defaultHeader),
      wrap(footer, defaultFooter)
    )
}

case class Margins private(
                            top: Option[Double],
                            bottom: Option[Double],
                            right: Option[Double],
                            left: Option[Double],
                            header: Option[Double],
                            footer: Option[Double]) {

  def withTop(top: Double) =
    copy(top = Option(top))

  def withBottom(bottom: Double) =
    copy(bottom = Option(bottom))

  def withRight(right: Double) =
    copy(right = Option(right))

  def withLeft(left: Double) =
    copy(left = Option(left))

  def withHeader(header: Double) =
    copy(header = Option(header))

  def withFooter(footer: Double) =
    copy(footer = Option(footer))

  def applyTo(sheet : XSSFSheet) {
    top.foreach(topMargin => sheet.setMargin(TopMargin, topMargin))
    bottom.foreach(bottomMargin => sheet.setMargin(BottomMargin, bottomMargin))
    right.foreach(rightMargin => sheet.setMargin(RightMargin, rightMargin))
    left.foreach(leftMargin => sheet.setMargin(LeftMargin, leftMargin))
    header.foreach(headerMargin => sheet.setMargin(HeaderMargin, headerMargin))
    footer.foreach(footerMargin => sheet.setMargin(FooterMargin, footerMargin))
  }
}