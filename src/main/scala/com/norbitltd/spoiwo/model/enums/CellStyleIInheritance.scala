package com.norbitltd.spoiwo.model.enums

object CellStyleInheritance {
  val CellOnly = CellStyleInheritance("CellOnly")
  val CellThenRow = CellStyleInheritance("CellThenRow")
  val CellThenColumn = CellStyleInheritance("CellThenColumn")
  val CellThenRowThenColumn = CellStyleInheritance("CellThenRowThenColumn")
  val CellThenColumnThenRow = CellStyleInheritance("CellThenColumnThenRow")
}

case class CellStyleInheritance private (value : String) {

  override def toString = value

}
