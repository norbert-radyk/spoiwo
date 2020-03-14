package com.norbitltd.spoiwo.model.enums

object CellFill {

  lazy val None: CellFill = CellFill("None")
  lazy val Solid: CellFill = CellFill("Solid")

  object Pattern {

    lazy val AltBars: CellFill = CellFill("AltBars")
    lazy val BigSpots: CellFill = CellFill("BigSpots")
    lazy val Bricks: CellFill = CellFill("Bricks")
    lazy val Diamonds: CellFill = CellFill("Diamonds")
    lazy val Squares: CellFill = CellFill("Squares")

    object Dots {
      lazy val Fine: CellFill = CellFill("Fine")
      lazy val Least: CellFill = CellFill("Least")
      lazy val Less: CellFill = CellFill("Less")
      lazy val Sparse: CellFill = CellFill("Sparse")
    }

    object Diagonals {
      lazy val ThinBackward: CellFill = CellFill("ThinBackward")
      lazy val ThinForward: CellFill = CellFill("ThinForward")
      lazy val ThickBackward: CellFill = CellFill("ThickBackward")
      lazy val ThickForward: CellFill = CellFill("ThickForward")
    }

    object Bands {
      lazy val ThickHorizontal: CellFill = CellFill("ThickHorizontal")
      lazy val ThickVertical: CellFill = CellFill("ThickVertical")
      lazy val ThinHorizontal: CellFill = CellFill("ThinHorizontal")
      lazy val ThinVertical: CellFill = CellFill("ThinVertical")
    }

  }

}

case class CellFill private (value: String) {

  override def toString: String = value

}
