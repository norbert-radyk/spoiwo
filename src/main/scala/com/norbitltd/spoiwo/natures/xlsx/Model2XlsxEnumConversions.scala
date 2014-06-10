package com.norbitltd.spoiwo.natures.xlsx

import org.apache.poi.ss.usermodel
import org.apache.poi.ss.usermodel._
import com.norbitltd.spoiwo.model.enums._
import com.norbitltd.spoiwo.model.enums.FontFamily
import com.norbitltd.spoiwo.model.enums.FontScheme
import com.norbitltd.spoiwo.model.enums.PageOrder
import com.norbitltd.spoiwo.model.enums.PaperSize

object Model2XlsxEnumConversions {

  def convertBorderStyle(borderStyle: CellBorderStyle): BorderStyle = {
    import CellBorderStyle._
    import BorderStyle._

    borderStyle match {
      case DashDot => DASH_DOT
      case DashDotDot => DASH_DOT_DOT
      case Dashed => DASHED
      case Dotted => DOTTED
      case Double => DOUBLE
      case Hair => HAIR
      case Medium => MEDIUM
      case MediumDashDot => MEDIUM_DASH_DOT
      case MediumDashDotDot => MEDIUM_DASH_DOT_DOTC
      case MediumDashed => MEDIUM_DASHED
      case CellBorderStyle.None => NONE
      case SlantedDashDot => SLANTED_DASH_DOT
      case Thick => THICK
      case Thin => THIN
      case CellBorderStyle(value) =>
        throw new IllegalArgumentException(s"Unable to convert BorderStyle=$value to XLSX - unsupported enum!")
    }
  }

  def convertBorderStyleFormatting(borderStyle: CellBorderStyle): Short = {
    import CellBorderStyle._
    import BorderFormatting._

    borderStyle match {
      case DashDot => BORDER_DASH_DOT
      case DashDotDot => BORDER_DASH_DOT_DOT
      case Dashed => BORDER_DASHED
      case Dotted => BORDER_DOTTED
      case Double => BORDER_DOUBLE
      case Hair => BORDER_HAIR
      case Medium => BORDER_MEDIUM
      case MediumDashDot => BORDER_MEDIUM_DASH_DOT
      case MediumDashDotDot => BORDER_MEDIUM_DASH_DOT_DOT
      case MediumDashed => BORDER_MEDIUM_DASHED
      case CellBorderStyle.None => BORDER_NONE
      case SlantedDashDot => BORDER_SLANTED_DASH_DOT
      case Thick => BORDER_THICK
      case Thin => BORDER_THIN
      case CellBorderStyle(value) =>
        throw new IllegalArgumentException(s"Unable to convert BorderStyle=$value to XLSX - unsupported enum!")
    }
  }

  def convertCellFill(cf: CellFill): FillPatternType = {
    import CellFill._

    cf match {
      case CellFill.None => FillPatternType.NO_FILL
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

  def convertCharset(charset: Charset): usermodel.FontCharset = {
    import Charset._
    import usermodel.FontCharset._

    charset match {
      case Charset.ANSI => FontCharset.ANSI
      case Arabic => ARABIC
      case Baltic => BALTIC
      case ChineseBig5 => CHINESEBIG5
      case Default => DEFAULT
      case EastEurope => EASTEUROPE
      case Charset.GB2312 => FontCharset.GB2312
      case Greek => GREEK
      case Hangeul => HANGEUL
      case Hebrew => HEBREW
      case Johab => JOHAB
      case Mac => MAC
      case Charset.OEM => FontCharset.OEM
      case Russian => RUSSIAN
      case ShiftJIS => SHIFTJIS
      case Symbol => SYMBOL
      case Thai => THAI
      case Turkish => TURKISH
      case Vietnamese => VIETNAMESE
      case Charset(value) =>
        throw new Exception(s"Unable to convert Charset=$value to XLSX - unsupported enum!")
    }
  }

  def convertFontFamily(fontFamily: FontFamily): usermodel.FontFamily = {
    import FontFamily._
    import usermodel.FontFamily._

    fontFamily match {
      case NotApplicable => NOT_APPLICABLE
      case Roman => ROMAN
      case Swiss => SWISS
      case Modern => MODERN
      case Script => SCRIPT
      case Decorative => DECORATIVE
      case FontFamily(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert FontFamily=$value to XLSX - unsupported enum!")

    }
  }

  def convertFontScheme(fontScheme: FontScheme): usermodel.FontScheme = {
    import FontScheme._
    import usermodel.FontScheme._

    fontScheme match {
      case FontScheme.None => NONE
      case Major => MAJOR
      case Minor => MINOR
      case FontScheme(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert FontScheme=$value to XLSX - unsupported enum!")
    }
  }

  def convertHorizontalAlignment(horizontalAlignment: CellHorizontalAlignment): HorizontalAlignment = {
    import CellHorizontalAlignment._
    import HorizontalAlignment._

    horizontalAlignment match {
      case Center => CENTER
      case CenterSelection => CENTER_SELECTION
      case Disturbed => DISTRIBUTED
      case Fill => FILL
      case General => GENERAL
      case Justify => JUSTIFY
      case Left => LEFT
      case Right => RIGHT
      case CellHorizontalAlignment(value) =>
        throw new IllegalArgumentException(s"Unable to convert HorizontalAlignment=$value to XLSX - unsupported enum!")
    }
  }

  def convertMissingCellPolicy(missingCellPolicy: MissingCellPolicy): usermodel.Row.MissingCellPolicy = {
    import MissingCellPolicy._

    missingCellPolicy match {
      case ReturnNullAndBlank => usermodel.Row.RETURN_NULL_AND_BLANK
      case ReturnBlankAsNull => usermodel.Row.RETURN_BLANK_AS_NULL
      case CreateNullAsBlank => usermodel.Row.CREATE_NULL_AS_BLANK
      case MissingCellPolicy(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert MissingCellPolicy=$value to XLSX - unsupported enum!")
    }
  }

  def convertPageOrder(po: PageOrder): usermodel.PageOrder = {
    import PageOrder._

    po match {
      case DownThenOver => usermodel.PageOrder.DOWN_THEN_OVER
      case OverThenDown => usermodel.PageOrder.OVER_THEN_DOWN
      case PageOrder(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert PageOrder=$value to XLSX - unsupported enum!")
    }
  }

  def convertPaperSize(ps: PaperSize): usermodel.PaperSize = {
    import PaperSize._

    ps match {
      case Letter => usermodel.PaperSize.LETTER_PAPER
      case LetterSmall => usermodel.PaperSize.LETTER_SMALL_PAPER
      case Tabloid => usermodel.PaperSize.TABLOID_PAPER
      case Ledger => usermodel.PaperSize.LEDGER_PAPER
      case Legal => usermodel.PaperSize.LEGAL_PAPER
      case Statement => usermodel.PaperSize.STATEMENT_PAPER
      case Executive => usermodel.PaperSize.EXECUTIVE_PAPER
      case A3 => usermodel.PaperSize.A3_PAPER
      case A4 => usermodel.PaperSize.A4_PAPER
      case A4Small => usermodel.PaperSize.A4_SMALL_PAPER
      case A5 => usermodel.PaperSize.A5_PAPER
      case B4 => usermodel.PaperSize.B4_PAPER
      case B5 => usermodel.PaperSize.B5_PAPER
      case Folio => usermodel.PaperSize.FOLIO_PAPER
      case Quarto => usermodel.PaperSize.QUARTO_PAPER
      case Standard10x14 => usermodel.PaperSize.STANDARD_PAPER_10_14
      case Standard11x17 => usermodel.PaperSize.STANDARD_PAPER_11_17
      case PaperSize(value) =>
        throw new IllegalArgumentException(s"Unable to convert PaperSize=$value to XLSX - unsupported enum!")
    }
  }

  def convertTypeOffset(typeOffset: TypeOffset): Short = {
    import TypeOffset._

    typeOffset match {
      case TypeOffset.None => usermodel.Font.SS_NONE
      case Subscript => usermodel.Font.SS_SUB
      case Superscript => usermodel.Font.SS_SUPER
      case TypeOffset(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert Type Offset = $value to XLSX - unsupported enum!")
    }
  }

  def convertEscapementType(typeOffset: TypeOffset): Short = {
    import TypeOffset._

    typeOffset match {
      case TypeOffset.None => FontFormatting.SS_NONE
      case Subscript => FontFormatting.SS_SUB
      case Superscript => FontFormatting.SS_SUPER
      case TypeOffset(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert Type Offset = $value to XLSX - unsupported enum!")
    }
  }

  def convertUnderline(underline: Underline): usermodel.FontUnderline = {
    import Underline._
    import usermodel.FontUnderline._

    underline match {
      case Double => DOUBLE
      case DoubleAccounting => DOUBLE_ACCOUNTING
      case Underline.None => NONE
      case Single => SINGLE
      case SingleAccounting => SINGLE_ACCOUNTING
      case Underline(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert Underline=$value to XLSX - unsupported enum!")
    }
  }

  def convertUnderlineFormatting(underline: Underline): Byte = {
    import Underline._
    import FontFormatting._

    underline match {
      case Double => U_DOUBLE
      case DoubleAccounting => U_DOUBLE_ACCOUNTING
      case Underline.None => U_NONE
      case Single => U_SINGLE
      case SingleAccounting => U_SINGLE_ACCOUNTING
      case Underline(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert Underline=$value to XLSX - unsupported enum!")
    }
  }

  def convertVerticalAlignment(verticalAlignment: CellVerticalAlignment): VerticalAlignment = {
    import CellVerticalAlignment._
    import VerticalAlignment._

    verticalAlignment match {
      case Bottom => BOTTOM
      case Center => CENTER
      case Disturbed => DISTRIBUTED
      case Justify => JUSTIFY
      case Top => TOP
      case CellVerticalAlignment(value) =>
        throw new IllegalArgumentException(s"Unable to convert VerticalAlignment=$value to XLSX - unsupported enum!")
    }
  }
}