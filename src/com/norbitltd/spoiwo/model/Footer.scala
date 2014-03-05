package com.norbitltd.spoiwo.model

object Footer extends Factory {

  private lazy val defaultLeft = defaultPOISheet.getFooter.getLeft
  private lazy val defaultCenter = defaultPOISheet.getFooter.getCenter
  private lazy val defaultRight = defaultPOISheet.getFooter.getRight
  private lazy val defaultFirstLeft = defaultPOISheet.getFirstFooter.getLeft
  private lazy val defaultFirstCenter = defaultPOISheet.getFirstFooter.getCenter
  private lazy val defaultFirstRight = defaultPOISheet.getFirstFooter.getRight
  private lazy val defaultOddLeft = defaultPOISheet.getOddFooter.getLeft
  private lazy val defaultOddCenter = defaultPOISheet.getOddFooter.getCenter
  private lazy val defaultOddRight = defaultPOISheet.getOddFooter.getRight
  private lazy val defaultEvenLeft = defaultPOISheet.getEvenFooter.getLeft
  private lazy val defaultEvenCenter = defaultPOISheet.getEvenFooter.getCenter
  private lazy val defaultEvenRight = defaultPOISheet.getEvenFooter.getRight

  val Empty = Standard()

  def Standard(left: String = defaultLeft,
            center: String = defaultCenter,
            right: String = defaultRight,
            firstLeft: String = defaultFirstLeft,
            firstCenter: String = defaultFirstCenter,
            firstRight: String = defaultFirstRight): Footer =
    Footer(left = wrap(left, defaultLeft),
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
    firstRight: String = defaultFirstRight) : Footer =
    Footer(
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

case class Footer private(
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

}