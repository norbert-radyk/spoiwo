package com.norbitltd.spoiwo.model

object CellFill {

  lazy val None = CellFill(0)
  lazy val Solid = CellFill(1)

  object Pattern {

    lazy val AltBars = CellFill(2)
    lazy val BigSpots = CellFill(3)
    lazy val Bricks = CellFill(4)
    lazy val Diamonds = CellFill(5)
    lazy val Squares = CellFill(6)

    object Dots {
      lazy val Fine = CellFill(10)
      lazy val Least = CellFill(11)
      lazy val Less = CellFill(12)
      lazy val Sparse = CellFill(13)
    }

    object Diagonals {
      lazy val ThinBackward = CellFill(20)
      lazy val ThinForward = CellFill(21)
      lazy val ThickBackward = CellFill(22)
      lazy val ThickForward = CellFill(23)
    }

    object Bands {
      lazy val ThickHorizontal = CellFill(30)
      lazy val ThickVertical = CellFill(31)
      lazy val ThinHorizontal = CellFill(32)
      lazy val ThinVertical = CellFill(33)
    }

  }

}

case class CellFill private(id: Int)
