package com.norbitltd.spoiwo.model

object MeasureUnit {
  lazy val Point = MeasureUnit("Point")
  lazy val Unit = MeasureUnit("Unit")
}

case class MeasureUnit private(value: String) {
  override def toString = value
}

object Measure {

  private[model] val Undefined = new Measure(-1, MeasureUnit.Point)

  implicit class MeasureEnrichment(value: Int) {
    def points = new Measure(value, MeasureUnit.Point)
    def units = new Measure(value, MeasureUnit.Unit)
  }
}

class Measure(measureValue: Int, measureUnit: MeasureUnit) {

  private val measureInUnits = measureUnit match {
    case MeasureUnit.Point => measureValue * 20
    case MeasureUnit.Unit => measureValue
    case _ => throw new IllegalArgumentException(
      s"Unable to convert Measure Unit = $measureUnit to XLSX - unsupported enum!")
  }

  def toUnits = measureInUnits

  def toPoints : Short = (measureInUnits / 20).toShort

}
