package com.norbitltd.spoiwo.model.enums

object CellStyleInheritance {
  val CellOnly: CellStyleInheritance = CellStyleInheritance("CellOnly")
  val CellThenRow: CellStyleInheritance = CellStyleInheritance("CellThenRow")
  val CellThenColumn: CellStyleInheritance = CellStyleInheritance("CellThenColumn")
  val CellThenSheet: CellStyleInheritance = CellStyleInheritance("CellThenSheet")
  val CellThenRowThenSheet: CellStyleInheritance = CellStyleInheritance("CellThenRowThenSheet")
  val CellThenColumnThenSheet: CellStyleInheritance = CellStyleInheritance("CellThenColumnThenSheet")
  val CellThenRowThenColumn: CellStyleInheritance = CellStyleInheritance("CellThenRowThenColumn")
  val CellThenColumnThenRow: CellStyleInheritance = CellStyleInheritance("CellThenColumnThenRow")
  val CellThenRowThenColumnThenSheet: CellStyleInheritance = CellStyleInheritance("CellThenRowThenColumnThenSheet")
  val CellThenColumnThenRowThenSheet: CellStyleInheritance = CellStyleInheritance("CellThenColumnThenRowThenSheet")
}

case class CellStyleInheritance private (value: String) {

  override def toString: String = value

}
