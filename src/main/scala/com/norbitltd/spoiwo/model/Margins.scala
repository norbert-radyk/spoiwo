package com.norbitltd.spoiwo.model

object Margins {

  def apply(top: java.lang.Double = null,
            bottom: java.lang.Double = null,
            right: java.lang.Double = null,
            left: java.lang.Double = null,
            header: java.lang.Double = null,
            footer: java.lang.Double = null): Margins =
    Margins(
      Option(top).map(_.doubleValue),
      Option(bottom).map(_.doubleValue),
      Option(right).map(_.doubleValue),
      Option(left).map(_.doubleValue),
      Option(header).map(_.doubleValue),
      Option(footer).map(_.doubleValue)
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

  def withoutTop =
    copy(top = None)

  def withBottom(bottom: Double) =
    copy(bottom = Option(bottom))

  def withoutBottom =
    copy(bottom = None)

  def withRight(right: Double) =
    copy(right = Option(right))

  def withoutRight =
    copy(right = None)

  def withLeft(left: Double) =
    copy(left = Option(left))

  def withoutLeft =
    copy(left = None)

  def withHeader(header: Double) =
    copy(header = Option(header))

  def withoutHeader =
    copy(header = None)

  def withFooter(footer: Double) =
    copy(footer = Option(footer))

  def withoutFooter =
    copy(footer = None)


}