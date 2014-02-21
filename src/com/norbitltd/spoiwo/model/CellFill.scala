package com.norbitltd.spoiwo.model

import org.apache.poi.ss.usermodel.FillPatternType

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

case class CellFill private(id: Int) {
  import CellFill._

  def convert(): FillPatternType = this match {
    case None => FillPatternType.NO_FILL
    case Solid => FillPatternType.SOLID_FOREGROUND
    case Pattern.AltBars => FillPatternType.ALT_BARS
    case Pattern.BigSpots => FillPatternType.BIG_SPOTS
    case Pattern.Bricks => FillPatternType.BRICKS
    case Pattern.Diamonds => FillPatternType.DIAMONDS
    case Pattern.Squares => FillPatternType.SQUARES
    case Pattern.Dots.Fine => FillPatternType.FINE_DOTS
    case Pattern.Dots.Least => FillPatternType.LEAST_DOTS
    case Pattern.Dots.Less => FillPatternType.LESS_DOTS
    case Pattern.Dots.Sparse => FillPatternType.SPARSE_DOTS
    case Pattern.Diagonals.ThickBackward => FillPatternType.THICK_BACKWARD_DIAG
    case Pattern.Diagonals.ThickForward => FillPatternType.THICK_FORWARD_DIAG
    case Pattern.Diagonals.ThinBackward => FillPatternType.THIN_BACKWARD_DIAG
    case Pattern.Diagonals.ThinForward => FillPatternType.THIN_FORWARD_DIAG
    case Pattern.Bands.ThickHorizontal => FillPatternType.THICK_HORZ_BANDS
    case Pattern.Bands.ThickVertical => FillPatternType.THICK_VERT_BANDS
    case Pattern.Bands.ThinHorizontal => FillPatternType.THIN_HORZ_BANDS
    case Pattern.Bands.ThinVertical => FillPatternType.THIN_VERT_BANDS
  }
}
