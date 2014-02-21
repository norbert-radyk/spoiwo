package com.norbitltd.spoiwo.model

object CellVerticalAlignment {

  lazy val Undefined = CellVerticalAlignment(0)
  lazy val Bottom = CellVerticalAlignment(1)
  lazy val Center = CellVerticalAlignment(2)
  lazy val Disturbed = CellVerticalAlignment(3)
  lazy val Justify = CellVerticalAlignment(4)
  lazy val Top = CellVerticalAlignment(5)

}

case class CellVerticalAlignment private(id : Int)


