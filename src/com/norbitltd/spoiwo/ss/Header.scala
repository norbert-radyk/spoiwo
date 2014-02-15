package com.norbitltd.spoiwo.ss

import org.apache.poi.xssf.usermodel.XSSFSheet

object Header extends Factory {

  private lazy val defaultLeft = defaultPOISheet.getHeader.getLeft
  private lazy val defaultCenter = defaultPOISheet.getHeader.getCenter
  private lazy val defaultRight = defaultPOISheet.getHeader.getRight
  private lazy val defaultFirstLeft = defaultPOISheet.getFirstHeader.getLeft
  private lazy val defaultFirstCenter = defaultPOISheet.getFirstHeader.getCenter
  private lazy val defaultFirstRight = defaultPOISheet.getFirstHeader.getRight
  private lazy val defaultOddLeft = defaultPOISheet.getOddHeader.getLeft
  private lazy val defaultOddCenter = defaultPOISheet.getOddHeader.getCenter
  private lazy val defaultOddRight = defaultPOISheet.getOddHeader.getRight
  private lazy val defaultEvenLeft = defaultPOISheet.getEvenHeader.getLeft
  private lazy val defaultEvenCenter = defaultPOISheet.getEvenHeader.getCenter
  private lazy val defaultEvenRight = defaultPOISheet.getEvenHeader.getRight

  val Empty = Standard()

  def Standard(left: String = defaultLeft,
               center: String = defaultCenter,
               right: String = defaultRight,
               firstLeft: String = defaultFirstLeft,
               firstCenter: String = defaultFirstCenter,
               firstRight: String = defaultFirstRight): Header =
    Header(left = wrap(left, defaultLeft),
      center = wrap(center, defaultCenter),
      right = wrap(right, defaultRight),
      firstLeft = wrap(firstLeft, defaultFirstLeft),
      firstCenter = wrap(firstCenter, defaultFirstCenter),
      firstRight = wrap(firstRight, defaultFirstRight),
      oddLeft = None, oddCenter = None, oddRight = None,
      evenLeft = None, evenCenter = None, evenRight = None)

  def EvenOdd(oddLeft : String = defaultOddLeft,
              oddCenter : String = defaultOddCenter,
              oddRight : String = defaultOddRight,
              evenLeft : String = defaultEvenLeft,
              evenCenter : String = defaultEvenCenter,
              evenRight : String = defaultEvenRight,
              firstLeft: String = defaultFirstLeft,
              firstCenter: String = defaultFirstCenter,
              firstRight: String = defaultFirstRight) : Header =
    Header(
      oddLeft = wrap(oddLeft, defaultOddLeft),
      oddCenter = wrap(oddCenter, defaultOddCenter),
      oddRight = wrap(oddRight, defaultOddRight),
      evenLeft = wrap(evenLeft, defaultEvenLeft),
      evenCenter = wrap(evenCenter, defaultEvenCenter),
      evenRight = wrap(evenRight, defaultEvenRight),
      firstLeft = wrap(firstLeft, defaultFirstLeft),
      firstCenter = wrap(firstCenter, defaultFirstCenter),
      firstRight = wrap(firstRight, defaultFirstRight),
      left = None, center = None, right = None
    )
}

case class Header private(
                           left: Option[String],
                           center: Option[String],
                           right: Option[String],
                           firstLeft: Option[String],
                           firstCenter: Option[String],
                           firstRight: Option[String],
                           oddLeft: Option[String],
                           oddCenter: Option[String],
                           oddRight: Option[String],
                           evenLeft: Option[String],
                           evenCenter: Option[String],
                           evenRight: Option[String]) {

  def withFirstPageLeft(firstLeft : String) =
    copy(firstLeft = Option(firstLeft))

  def withFirstPageCenter(firstCenter : String) =
    copy(firstCenter = Option(firstCenter))

  def withFirstPageRight(firstRight : String) =
    copy(firstLeft = Option(firstRight))

  def applyTo(sheet: XSSFSheet) {
    left.foreach(sheet.getHeader.setLeft)
    center.foreach(sheet.getHeader.setCenter)
    right.foreach(sheet.getHeader.setRight)

    firstLeft.foreach(sheet.getFirstHeader.setLeft)
    firstCenter.foreach(sheet.getFirstHeader.setCenter)
    firstRight.foreach(sheet.getFirstHeader.setRight)

    oddLeft.foreach(sheet.getOddHeader.setLeft)
    oddCenter.foreach(sheet.getOddHeader.setCenter)
    oddRight.foreach(sheet.getOddHeader.setRight)

    evenLeft.foreach(sheet.getEvenHeader.setLeft)
    evenCenter.foreach(sheet.getEvenHeader.setCenter)
    evenRight.foreach(sheet.getEvenHeader.setRight)
  }
}