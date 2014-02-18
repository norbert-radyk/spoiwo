package com.norbitltd.spoiwo.model

import org.apache.poi.ss.usermodel.FillPatternType

object Fill {

  lazy val None = Fill(0)
  lazy val Solid = Fill(1)

  object Pattern {

    lazy val AltBars = Fill(2)
    lazy val BigSpots = Fill(3)
    lazy val Bricks = Fill(4)
    lazy val Diamonds = Fill(5)
    lazy val Squares = Fill(6)

    object Dots {
      lazy val Fine = Fill(10)
      lazy val Least = Fill(11)
      lazy val Less = Fill(12)
      lazy val Sparse = Fill(13)
    }

    object Diagonals {
      lazy val ThinBackward = Fill(20)
      lazy val ThinForward = Fill(21)
      lazy val ThickBackward = Fill(22)
      lazy val ThickForward = Fill(23)
    }

    object Bands {
      lazy val ThickHorizontal = Fill(30)
      lazy val ThickVertical = Fill(31)
      lazy val ThinHorizontal = Fill(32)
      lazy val ThinVertical = Fill(33)
    }

  }





}


case class Fill private(id: Int) {
  import Fill._

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
