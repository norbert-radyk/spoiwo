package com.norbitltd.spoiwo.model.enums

object CellFill {

  lazy val None = CellFill("None")
  lazy val Solid = CellFill("Solid")

  object Pattern {

    lazy val AltBars = CellFill("AltBars")
    lazy val BigSpots = CellFill("BigSpots")
    lazy val Bricks = CellFill("Bricks")
    lazy val Diamonds = CellFill("Diamonds")
    lazy val Squares = CellFill("Squares")

    object Dots {
      lazy val Fine = CellFill("Fine")
      lazy val Least = CellFill("Least")
      lazy val Less = CellFill("Less")
      lazy val Sparse = CellFill("Sparse")
    }

    object Diagonals {
      lazy val ThinBackward = CellFill("ThinBackward")
      lazy val ThinForward = CellFill("ThinForward")
      lazy val ThickBackward = CellFill("ThickBackward")
      lazy val ThickForward = CellFill("ThickForward")
    }

    object Bands {
      lazy val ThickHorizontal = CellFill("ThickHorizontal")
      lazy val ThickVertical = CellFill("ThickVertical")
      lazy val ThinHorizontal = CellFill("ThinHorizontal")
      lazy val ThinVertical = CellFill("ThinVertical")
    }

  }

}

case class CellFill private (value: String) {

  override def toString: String = value

}
