package com.norbitltd.spoiwo.model

object TableStyle {

  def apply(
      name: TableStyleName,
      showColumnStripes: java.lang.Boolean = null,
      showRowStripes: java.lang.Boolean = null
  ): TableStyle =
    TableStyle(
      name = name,
      showColumnStripes = Option(showColumnStripes).map(_.booleanValue),
      showRowStripes = Option(showRowStripes).map(_.booleanValue)
    )
}

case class TableStyle private (
    name: TableStyleName,
    showColumnStripes: Option[Boolean],
    showRowStripes: Option[Boolean]
) {

  def withName(name: TableStyleName): TableStyle =
    copy(name = name)

  def withColumnStripes: TableStyle =
    copy(showColumnStripes = Some(true))

  def withoutColumnStripes: TableStyle =
    copy(showColumnStripes = Some(false))

  def withRowStripes: TableStyle =
    copy(showRowStripes = Some(true))

  def withoutRowStripes: TableStyle =
    copy(showRowStripes = Some(false))
}
