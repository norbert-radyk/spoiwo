package com.norbitltd.spoiwo.model

object TableStyle {

  def apply(
    name: TableStyleName,
    showColumnStripes: java.lang.Boolean = null,
    showRowStripes: java.lang.Boolean = null
  ): TableStyle =
    TableStyle(
      name              = name,
      showColumnStripes = Option(showColumnStripes).map(_.booleanValue),
      showRowStripes    = Option(showRowStripes).map(_.booleanValue)
    )
}

case class TableStyle private(
  name: TableStyleName,
  showColumnStripes: Option[Boolean],
  showRowStripes: Option[Boolean]
) {

  def withName(name: TableStyleName) =
    copy(name = name)

  def withColumnStripes =
    copy(showColumnStripes = Some(true))

  def withoutColumnStripes =
    copy(showColumnStripes = Some(false))

  def withRowStripes =
    copy(showRowStripes = Some(true))

  def withoutRowStripes =
    copy(showRowStripes = Some(false))
}
