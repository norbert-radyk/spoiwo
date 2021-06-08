package spoiwo.natures.xlsx

import org.apache.poi.common.usermodel.fonts.FontCharset
import org.apache.poi.ss.usermodel
import org.apache.poi.ss.usermodel._
import spoiwo.model.enums
import spoiwo.model.enums.CellFill.Pattern
import spoiwo.model.enums._

object Model2XlsxEnumConversions {

  def convertBorderStyle(borderStyle: CellBorderStyle): BorderStyle = {
    import CellBorderStyle._
    import BorderStyle._

    borderStyle match {
      case DashDot              => DASH_DOT
      case DashDotDot           => DASH_DOT_DOT
      case Dashed               => DASHED
      case Dotted               => DOTTED
      case Double               => DOUBLE
      case Hair                 => HAIR
      case Medium               => MEDIUM
      case MediumDashDot        => MEDIUM_DASH_DOT
      case MediumDashed         => MEDIUM_DASHED
      case CellBorderStyle.None => NONE
      case SlantedDashDot       => SLANTED_DASH_DOT
      case Thick                => THICK
      case Thin                 => THIN
      case CellBorderStyle(value) =>
        throw new IllegalArgumentException(s"Unable to convert BorderStyle=$value to XLSX - unsupported enum!")
    }
  }

  def convertCellFill(cf: CellFill): FillPatternType = {
    cf match {
      case CellFill.None                   => FillPatternType.NO_FILL
      case CellFill.Solid                  => FillPatternType.SOLID_FOREGROUND
      case Pattern.AltBars                 => FillPatternType.ALT_BARS
      case Pattern.BigSpots                => FillPatternType.BIG_SPOTS
      case Pattern.Bricks                  => FillPatternType.BRICKS
      case Pattern.Diamonds                => FillPatternType.DIAMONDS
      case Pattern.Squares                 => FillPatternType.SQUARES
      case Pattern.Dots.Fine               => FillPatternType.FINE_DOTS
      case Pattern.Dots.Least              => FillPatternType.LEAST_DOTS
      case Pattern.Dots.Less               => FillPatternType.LESS_DOTS
      case Pattern.Dots.Sparse             => FillPatternType.SPARSE_DOTS
      case Pattern.Diagonals.ThickBackward => FillPatternType.THICK_BACKWARD_DIAG
      case Pattern.Diagonals.ThickForward  => FillPatternType.THICK_FORWARD_DIAG
      case Pattern.Diagonals.ThinBackward  => FillPatternType.THIN_BACKWARD_DIAG
      case Pattern.Diagonals.ThinForward   => FillPatternType.THIN_FORWARD_DIAG
      case Pattern.Bands.ThickHorizontal   => FillPatternType.THICK_HORZ_BANDS
      case Pattern.Bands.ThickVertical     => FillPatternType.THICK_VERT_BANDS
      case Pattern.Bands.ThinHorizontal    => FillPatternType.THIN_HORZ_BANDS
      case Pattern.Bands.ThinVertical      => FillPatternType.THIN_VERT_BANDS
      case unexpected =>
        throw new IllegalArgumentException(s"Unable to convert CellFill=$unexpected to XLSX - unsupported enum!")
    }
  }

  def convertCharset(charset: Charset): FontCharset = {
    import FontCharset._
    import Charset._

    charset match {
      case Charset.ANSI   => FontCharset.ANSI
      case Arabic         => ARABIC
      case Baltic         => BALTIC
      case ChineseBig5    => CHINESEBIG5
      case Default        => DEFAULT
      case EastEurope     => EASTEUROPE
      case Charset.GB2312 => FontCharset.GB2312
      case Greek          => GREEK
      case Hangeul        => HANGUL
      case Hebrew         => HEBREW
      case Johab          => JOHAB
      case Mac            => MAC
      case Charset.OEM    => FontCharset.OEM
      case Russian        => RUSSIAN
      case ShiftJIS       => SHIFTJIS
      case Symbol         => SYMBOL
      case Thai           => THAI
      case Turkish        => TURKISH
      case Vietnamese     => VIETNAMESE
      case Charset(value) =>
        throw new Exception(s"Unable to convert Charset=$value to XLSX - unsupported enum!")
    }
  }

  def convertFontFamily(fontFamily: spoiwo.model.enums.FontFamily): usermodel.FontFamily = {
    import spoiwo.model.enums.FontFamily._
    import usermodel.FontFamily._

    fontFamily match {
      case NotApplicable => NOT_APPLICABLE
      case Roman         => ROMAN
      case Swiss         => SWISS
      case Modern        => MODERN
      case Script        => SCRIPT
      case Decorative    => DECORATIVE
      case spoiwo.model.enums.FontFamily(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert FontFamily=$value to XLSX - unsupported enum!")

    }
  }

  def convertFontScheme(fontScheme: spoiwo.model.enums.FontScheme): usermodel.FontScheme = {
    import spoiwo.model.enums.FontScheme._
    import usermodel.FontScheme._

    fontScheme match {
      case spoiwo.model.enums.FontScheme.None => NONE
      case Major                              => MAJOR
      case Minor                              => MINOR
      case spoiwo.model.enums.FontScheme(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert FontScheme=$value to XLSX - unsupported enum!")
    }
  }

  def convertHyperlinkType(hyperlinkType: HyperLinkType): org.apache.poi.common.usermodel.HyperlinkType = {
    import org.apache.poi.common.usermodel.HyperlinkType._
    import spoiwo.model.enums.HyperLinkType._
    hyperlinkType match {
      case Url      => URL
      case Document => DOCUMENT
      case Email    => EMAIL
      case File     => FILE
      case spoiwo.model.enums.HyperLinkType(value) =>
        throw new IllegalArgumentException(s"Unable to convert HyperlinkType=$value to XLSX - unsupported enum!")
    }
  }

  def convertHorizontalAlignment(horizontalAlignment: CellHorizontalAlignment): HorizontalAlignment = {
    import HorizontalAlignment._
    import spoiwo.model.enums.CellHorizontalAlignment._

    horizontalAlignment match {
      case Center          => CENTER
      case CenterSelection => CENTER_SELECTION
      case Disturbed       => DISTRIBUTED
      case Fill            => FILL
      case General         => GENERAL
      case Justify         => JUSTIFY
      case Left            => LEFT
      case Right           => RIGHT
      case spoiwo.model.enums.CellHorizontalAlignment(value) =>
        throw new IllegalArgumentException(s"Unable to convert HorizontalAlignment=$value to XLSX - unsupported enum!")
    }
  }

  def convertMissingCellPolicy(missingCellPolicy: MissingCellPolicy): usermodel.Row.MissingCellPolicy = {
    import spoiwo.model.enums.MissingCellPolicy._
    missingCellPolicy match {
      case ReturnNullAndBlank => usermodel.Row.MissingCellPolicy.RETURN_NULL_AND_BLANK
      case ReturnBlankAsNull  => usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL
      case CreateNullAsBlank  => usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK
      case spoiwo.model.enums.MissingCellPolicy(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert MissingCellPolicy=$value to XLSX - unsupported enum!")
    }
  }

  def convertPageOrder(po: spoiwo.model.enums.PageOrder): usermodel.PageOrder = {
    import spoiwo.model.enums.PageOrder._

    po match {
      case DownThenOver => usermodel.PageOrder.DOWN_THEN_OVER
      case OverThenDown => usermodel.PageOrder.OVER_THEN_DOWN
      case spoiwo.model.enums.PageOrder(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert PageOrder=$value to XLSX - unsupported enum!")
    }
  }

  def convertPaperSize(ps: spoiwo.model.enums.PaperSize): usermodel.PaperSize = {
    import spoiwo.model.enums.PaperSize._

    ps match {
      case Letter        => usermodel.PaperSize.LETTER_PAPER
      case LetterSmall   => usermodel.PaperSize.LETTER_SMALL_PAPER
      case Tabloid       => usermodel.PaperSize.TABLOID_PAPER
      case Ledger        => usermodel.PaperSize.LEDGER_PAPER
      case Legal         => usermodel.PaperSize.LEGAL_PAPER
      case Statement     => usermodel.PaperSize.STATEMENT_PAPER
      case Executive     => usermodel.PaperSize.EXECUTIVE_PAPER
      case A3            => usermodel.PaperSize.A3_PAPER
      case A4            => usermodel.PaperSize.A4_PAPER
      case A4Small       => usermodel.PaperSize.A4_SMALL_PAPER
      case A5            => usermodel.PaperSize.A5_PAPER
      case B4            => usermodel.PaperSize.B4_PAPER
      case B5            => usermodel.PaperSize.B5_PAPER
      case Folio         => usermodel.PaperSize.FOLIO_PAPER
      case Quarto        => usermodel.PaperSize.QUARTO_PAPER
      case Standard10x14 => usermodel.PaperSize.STANDARD_PAPER_10_14
      case Standard11x17 => usermodel.PaperSize.STANDARD_PAPER_11_17
      case spoiwo.model.enums.PaperSize(value) =>
        throw new IllegalArgumentException(s"Unable to convert PaperSize=$value to XLSX - unsupported enum!")
    }
  }

  def convertReadingOrder(readingOrder: CellReadingOrder): ReadingOrder = {
    import ReadingOrder._
    import spoiwo.model.enums.CellReadingOrder._

    readingOrder match {
      case Context     => CONTEXT
      case LeftToRight => LEFT_TO_RIGHT
      case RightToLeft => RIGHT_TO_LEFT
      case spoiwo.model.enums.CellReadingOrder(value) =>
        throw new IllegalArgumentException(s"Unable to convert ReadingOrder=$value to XLSX - unsupported enum!")
    }
  }

  def convertTypeOffset(typeOffset: TypeOffset): Short = {
    import spoiwo.model.enums.TypeOffset._

    typeOffset match {
      case TypeOffset.None => usermodel.Font.SS_NONE
      case Subscript       => usermodel.Font.SS_SUB
      case Superscript     => usermodel.Font.SS_SUPER
      case spoiwo.model.enums.TypeOffset(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert Type Offset = $value to XLSX - unsupported enum!")
    }
  }

  def convertUnderline(underline: Underline): usermodel.FontUnderline = {
    import usermodel.FontUnderline._
    import spoiwo.model.enums.Underline._

    underline match {
      case Double           => DOUBLE
      case DoubleAccounting => DOUBLE_ACCOUNTING
      case Underline.None   => NONE
      case Single           => SINGLE
      case SingleAccounting => SINGLE_ACCOUNTING
      case spoiwo.model.enums.Underline(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert Underline=$value to XLSX - unsupported enum!")
    }
  }

  def convertUnderlineFormatting(underline: Underline): Byte = {
    import spoiwo.model.enums.Underline._
    underline match {
      case Double           => Font.U_DOUBLE
      case DoubleAccounting => Font.U_DOUBLE_ACCOUNTING
      case Underline.None   => Font.U_NONE
      case Single           => Font.U_SINGLE
      case SingleAccounting => Font.U_SINGLE_ACCOUNTING
      case spoiwo.model.enums.Underline(value: String) =>
        throw new IllegalArgumentException(s"Unable to convert Underline=$value to XLSX - unsupported enum!")
    }
  }

  def convertVerticalAlignment(verticalAlignment: CellVerticalAlignment): VerticalAlignment = {
    import VerticalAlignment._
    import spoiwo.model.enums.CellVerticalAlignment._
    verticalAlignment match {
      case Bottom    => BOTTOM
      case Center    => CENTER
      case Disturbed => DISTRIBUTED
      case Justify   => JUSTIFY
      case Top       => TOP
      case spoiwo.model.enums.CellVerticalAlignment(value) =>
        throw new IllegalArgumentException(s"Unable to convert VerticalAlignment=$value to XLSX - unsupported enum!")
    }
  }
}
