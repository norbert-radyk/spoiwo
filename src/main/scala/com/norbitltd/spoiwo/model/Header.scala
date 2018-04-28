package com.norbitltd.spoiwo.model

object Header {

  val Empty: Header = Standard()

  def Standard(left: String = null,
               center: String = null,
               right: String = null,
               firstLeft: String = null,
               firstCenter: String = null,
               firstRight: String = null): Header =
    Header(
      left = Option(left),
      center = Option(center),
      right = Option(right),
      firstLeft = Option(firstLeft),
      firstCenter = Option(firstCenter),
      firstRight = Option(firstRight),
      oddLeft = None,
      oddCenter = None,
      oddRight = None,
      evenLeft = None,
      evenCenter = None,
      evenRight = None
    )

  def EvenOdd(oddLeft: String = null,
              oddCenter: String = null,
              oddRight: String = null,
              evenLeft: String = null,
              evenCenter: String = null,
              evenRight: String = null,
              firstLeft: String = null,
              firstCenter: String = null,
              firstRight: String = null): Header =
    Header(
      oddLeft = Option(oddLeft),
      oddCenter = Option(oddCenter),
      oddRight = Option(oddRight),
      evenLeft = Option(evenLeft),
      evenCenter = Option(evenCenter),
      evenRight = Option(evenRight),
      firstLeft = Option(firstLeft),
      firstCenter = Option(firstCenter),
      firstRight = Option(firstRight),
      left = None,
      center = None,
      right = None
    )
}

case class Header private (left: Option[String],
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

  def withFirstPageLeft(firstLeft: String): Header =
    copy(firstLeft = Option(firstLeft))

  def withFirstPageCenter(firstCenter: String): Header =
    copy(firstCenter = Option(firstCenter))

  def withFirstPageRight(firstRight: String): Header =
    copy(firstLeft = Option(firstRight))

}
