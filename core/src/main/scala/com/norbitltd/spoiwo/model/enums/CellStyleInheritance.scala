package com.norbitltd.spoiwo.model.enums

object CellStyleInheritance {
  val CellOnly = CellStyleInheritance("CellOnly")
  val CellThenRow = CellStyleInheritance("CellThenRow")
  val CellThenColumn = CellStyleInheritance("CellThenColumn")
  val CellThenSheet = CellStyleInheritance("CellThenSheet")
  val CellThenRowThenSheet = CellStyleInheritance("CellThenRowThenSheet")
  val CellThenColumnThenSheet = CellStyleInheritance("CellThenColumnThenSheet")
  val CellThenRowThenColumn = CellStyleInheritance("CellThenRowThenColumn")
  val CellThenColumnThenRow = CellStyleInheritance("CellThenColumnThenRow")
  val CellThenRowThenColumnThenSheet = CellStyleInheritance("CellThenRowThenColumnThenSheet")
  val CellThenColumnThenRowThenSheet = CellStyleInheritance("CellThenColumnThenRowThenSheet")
}

case class CellStyleInheritance private (value: String) {

  override def toString: String = value

}
