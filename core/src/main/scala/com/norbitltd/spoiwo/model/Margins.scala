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

case class Margins private (top: Option[Double],
                            bottom: Option[Double],
                            right: Option[Double],
                            left: Option[Double],
                            header: Option[Double],
                            footer: Option[Double]) {

  def withTop(top: Double): Margins =
    copy(top = Option(top))

  def withoutTop: Margins =
    copy(top = None)

  def withBottom(bottom: Double): Margins =
    copy(bottom = Option(bottom))

  def withoutBottom: Margins =
    copy(bottom = None)

  def withRight(right: Double): Margins =
    copy(right = Option(right))

  def withoutRight: Margins =
    copy(right = None)

  def withLeft(left: Double): Margins =
    copy(left = Option(left))

  def withoutLeft: Margins =
    copy(left = None)

  def withHeader(header: Double): Margins =
    copy(header = Option(header))

  def withoutHeader: Margins =
    copy(header = None)

  def withFooter(footer: Double): Margins =
    copy(footer = Option(footer))

  def withoutFooter: Margins =
    copy(footer = None)

}
