package com.norbitltd.spoiwo.model

object HeightUnit {
  lazy val Point = HeightUnit("Point")
  lazy val Unit = HeightUnit("Unit")
}

case class HeightUnit private (value: String) {
  override def toString: String = value
}

object Height {

  private[model] val Undefined = new Height(-1, HeightUnit.Point)

  implicit class HeightEnrichment(value: Int) {
    def points = new Height(value, HeightUnit.Point)
    def unitsOfHeight = new Height(value, HeightUnit.Unit)
  }
}

class Height(measureValue: Int, measureUnit: HeightUnit) {

  private val heightInUnits = measureUnit match {
    case HeightUnit.Point => measureValue * 20
    case HeightUnit.Unit  => measureValue
    case _ =>
      throw new IllegalArgumentException(s"Unable to convert Height Unit = $measureUnit to XLSX - unsupported enum!")
  }

  def toUnits: Int = heightInUnits

  def toPoints: Short = (heightInUnits / 20).toShort

}
