package com.norbitltd.spoiwo.model

import java.util.{Calendar, Date}
import org.joda.time.{LocalDate, DateTime}

//TODO Customized Formula Error
import org.apache.poi.ss.usermodel.FormulaError

object Cell extends Factory {

  val defaultActive = false
  val defaultIndex = -1
  val defaultStyle = CellStyle.Default

  lazy val Empty = apply("")

  def apply(value: String): Cell = apply(value, defaultIndex, defaultStyle)

  def apply(value: String, index: Int): Cell = apply(value, index, defaultStyle)

  def apply(value: String, style: CellStyle): Cell = apply(value, defaultIndex, style)

  def apply(value: String, index: Int, style: CellStyle): Cell = {

    val indexOption = wrap(index, defaultIndex)
    val styleOption = wrap(style, defaultStyle)

    if (value.startsWith("=")) {
      FormulaCell(value, indexOption, styleOption)
    } else if (value.contains("\n")) {
      StringCell(value, indexOption, Option(style.withWrapText))
    } else {
      StringCell(value, indexOption, styleOption)
    }
  }

  def apply(value: Double): Cell = apply(value, defaultIndex, defaultStyle)

  def apply(value: Double, index: Int): Cell = apply(value, index, defaultStyle)

  def apply(value: Double, style: CellStyle): Cell = apply(value, defaultIndex, style)

  def apply(value: Double, index: Int, style: CellStyle): Cell =
    NumericCell(value, wrap(index, defaultIndex), wrap(style, defaultStyle))

  def apply(value: Int): Cell = apply(value, defaultIndex, defaultStyle)

  def apply(value: Int, index: Int): Cell = apply(value, index, defaultStyle)

  def apply(value: Int, style: CellStyle): Cell = apply(value, defaultIndex, style)

  def apply(value: Int, index: Int, style: CellStyle): Cell =
    apply(value.toDouble, index, style)

  def apply(value: Long): Cell = apply(value, defaultIndex, defaultStyle)

  def apply(value: Long, index: Int): Cell = apply(value, index, defaultStyle)

  def apply(value: Long, style: CellStyle): Cell = apply(value, defaultIndex, style)

  def apply(value: Long, index: Int, style: CellStyle): Cell =
    apply(value.toDouble, index, style)

  def apply(value: Boolean): Cell = apply(value, defaultIndex, defaultStyle)

  def apply(value: Boolean, index: Int): Cell = apply(value, index, defaultStyle)

  def apply(value: Boolean, style: CellStyle): Cell = apply(value, defaultIndex, style)

  def apply(value: Boolean, index: Int, style: CellStyle): Cell =
    BooleanCell(value, wrap(index, defaultIndex), wrap(style, defaultStyle))

  def apply(value: Date): Cell = apply(value, defaultIndex, defaultStyle)

  def apply(value: DateTime) : Cell = apply(value.toDate)

  def apply(value: LocalDate) : Cell = apply(value.toDate)

  def apply(value: Date, index: Int): Cell = apply(value, index, defaultStyle)

  def apply(value: DateTime, index: Int) : Cell = apply(value.toDate, index)

  def apply(value: LocalDate, index: Int) : Cell = apply(value.toDate, index)

  def apply(value: Date, style: CellStyle): Cell = apply(value, defaultIndex, style)

  def apply(value: DateTime, style: CellStyle) : Cell = apply(value.toDate, style)

  def apply(value: LocalDate, style: CellStyle) : Cell = apply(value.toDate, style)

  def apply(value: Date, index: Int, style: CellStyle): Cell =
    DateCell(value, wrap(index, defaultIndex), wrap(style, defaultStyle))

  def apply(value: DateTime, index: Int, style: CellStyle): Cell =
    apply(value.toDate, index, style)

  def apply(value: LocalDate, index: Int, style: CellStyle): Cell =
    apply(value.toDate, index, style)

  def apply(value: Calendar): Cell = apply(value, defaultIndex, defaultStyle)

  def apply(value: Calendar, index: Int): Cell = apply(value, index, defaultStyle)

  def apply(value: Calendar, style: CellStyle): Cell = apply(value, defaultIndex, style)

  def apply(value: Calendar, index: Int, style: CellStyle): Cell =
    CalendarCell(value, wrap(index, defaultIndex), wrap(style, defaultStyle))
}

sealed abstract class Cell(index: Option[Int], style: Option[CellStyle]) {

  def getIndex = index

  def getStyle = style

}

case class StringCell(value: String, index: Option[Int], style: Option[CellStyle])
  extends Cell(index, style)   {
}

case class FormulaCell(formula: String, index: Option[Int], style: Option[CellStyle])
  extends Cell(index, style)

case class ErrorValueCell(formulaError: FormulaError, index: Option[Int], style: Option[CellStyle])
  extends Cell(index, style)

case class NumericCell(value: Double, index: Option[Int], style: Option[CellStyle])
  extends Cell(index, style)

case class BooleanCell(value: Boolean, index: Option[Int], style: Option[CellStyle])
  extends Cell(index, style)

case class DateCell(value: Date, index: Option[Int], style: Option[CellStyle])
  extends Cell(index, style)

case class CalendarCell(value: Calendar, index: Option[Int], style: Option[CellStyle])
  extends Cell(index, style)