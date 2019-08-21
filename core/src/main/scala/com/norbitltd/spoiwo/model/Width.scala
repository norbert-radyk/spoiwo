package com.norbitltd.spoiwo.model

object WidthUnit {
  lazy val Character = WidthUnit("Character")
  lazy val Unit = WidthUnit("Unit")
}

case class WidthUnit private (value: String) {
  override def toString: String = value
}

object Width {

  private[model] val Undefined = new Width(-1, WidthUnit.Character)

  implicit class WidthEnrichment(value: Int) {
    def characters = new Width(value, WidthUnit.Character)
    def unitsOfWidth = new Width(value, WidthUnit.Unit)
  }
}

class Width(measureValue: Int, measureUnit: WidthUnit) {

  private val widthInUnits: Int = measureUnit match {
    case WidthUnit.Character => measureValue * 256
    case WidthUnit.Unit      => measureValue
    case _ =>
      throw new IllegalArgumentException(
        s"Unable to convert Width Unit = $measureUnit to XLSX - unsupported enum!"
      )
  }

  def toUnits: Int = widthInUnits

  def toCharacters: Short = (widthInUnits / 256).toShort

}
