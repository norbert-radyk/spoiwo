package com.norbitltd.spoiwo.natures.xlsx


import org.apache.poi.ss.usermodel.{HorizontalAlignment, VerticalAlignment, BorderStyle}
import com.norbitltd.spoiwo.model.{CellBorderStyle, CellHorizontalAlignment, CellVerticalAlignment}

object Model2XlsxConversions {

  implicit class XlsxBorderStyle(bs: CellBorderStyle) {
    def convertAsXlsx() = convertBorderStyle(bs)
  }

  implicit class XlsxHorizontalAlignment(ha : CellHorizontalAlignment) {
    def convertAsXlsx() = convertHorizontalAlignment(ha)
  }

  implicit class XlsxVerticalAlignment(va : CellVerticalAlignment) {
    def convertAsXlsx() = convertVerticalAlignment(va)
  }

  def convertBorderStyle(borderStyle : CellBorderStyle) : BorderStyle = {
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
      case None => NONE
      case SlantedDashDot => SLANTED_DASH_DOT
      case Thick => THICK
      case Thin => THIN
      case CellBorderStyle(id) =>
        throw new Exception("Unsupported option for XLSX conversion with id=%d".format(id))
    }
  }

  def convertHorizontalAlignment(horizontalAlignment : CellHorizontalAlignment) : HorizontalAlignment = {
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
      case Undefined =>
        throw new RuntimeException("Internal SPOIWO framework error. Undefined value is not allowed for explicit transformation!")
      case CellHorizontalAlignment(id) =>
        throw new Exception("Unsupported option for XLSX conversion with id=%d".format(id))
    }
  }

  private def convertVerticalAlignment(verticalAlignment : CellVerticalAlignment) : VerticalAlignment = {
    import CellVerticalAlignment._
    import VerticalAlignment._

    verticalAlignment match {
      case Bottom => BOTTOM
      case Center => CENTER
      case Disturbed => DISTRIBUTED
      case Justify => JUSTIFY
      case Top => TOP
      case Undefined =>
        throw new RuntimeException("Internal SPOIWO framework error. Undefined value is not allowed for explicit transformation!")
      case CellVerticalAlignment(id) =>
        throw new RuntimeException("Unsupported option for XLSX conversion with id=%d".format(id))
    }
  }

}
