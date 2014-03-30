package com.norbitltd.spoiwo.model

import java.util.{Calendar, Date}
import org.joda.time.{LocalDate, DateTime}

sealed class CellValueType[T]
object CellValueType {
  implicit object StringWitness extends CellValueType[String]
  implicit object DoubleWitness extends CellValueType[Double]
  implicit object IntWitness extends CellValueType[Int]
  implicit object LongWitness extends CellValueType[Long]
  implicit object BooleanWitness extends CellValueType[Boolean]
  implicit object DateWitness extends CellValueType[Date]
  implicit object DateTimeWitness extends CellValueType[DateTime]
  implicit object LocalDateWitness extends CellValueType[LocalDate]
  implicit object CalendarWitness extends CellValueType[Calendar]
}

object Cell {

  lazy val Empty = apply("")

  def apply[T : CellValueType](value : T, index : java.lang.Integer = null, style : CellStyle = null) : Cell = {
    val indexOption = Option(index).map(_.intValue)
    val styleOption = Option(style)
    value match {
      case v : String => if (v.startsWith("=")) {
        FormulaCell(v.drop(1), indexOption, styleOption)
      } else if (v.contains("\n")) {
        StringCell(v, indexOption, Option(style.withWrapText))
      } else {
        StringCell(v, indexOption, styleOption)
      }
      case v : Double => NumericCell(v, indexOption, styleOption)
      case v : Int => NumericCell(v.toDouble, indexOption, styleOption)
      case v : Long => NumericCell(v.toDouble, indexOption, styleOption)
      case v : Boolean => BooleanCell(v, indexOption, styleOption)
      case v : Date => DateCell(v, indexOption, styleOption)
      case v : DateTime => DateCell(v.toDate, indexOption, styleOption)
      case v : LocalDate => DateCell(v.toDate, indexOption, styleOption)
      case v : Calendar => CalendarCell(v, indexOption, styleOption)
    }
  }
}

sealed trait Cell {

  val value : Any
  val index : Option[Int]
  val style : Option[CellStyle]

  protected def copyCell(value : Any = value, index : Option[Int] = index, style : Option[CellStyle] = style) : Cell

  def withIndex(index : Int) =
    copyCell(index = Option(index))

  def withoutIndex =
    copyCell(index = None)

  def withStyle(style : CellStyle) =
    copyCell(style = Option(style))

  def withoutStyle =
    copyCell(style = None)

}

case class StringCell private[model](value: String, index: Option[Int], style: Option[CellStyle])
  extends Cell {
  def copyCell(value : Any = value, index : Option[Int] = index, style : Option[CellStyle] = style) : Cell =
    copy(value.asInstanceOf[String], index, style)
}

case class FormulaCell private[model](value: String,  index: Option[Int], style: Option[CellStyle])
  extends Cell {
  def copyCell(value : Any = value, index : Option[Int] = index, style : Option[CellStyle] = style) : Cell =
    copy(value.asInstanceOf[String], index, style)
}

case class NumericCell private[model](value: Double, index: Option[Int], style: Option[CellStyle])
  extends Cell {
  def copyCell(value : Any = value, index : Option[Int] = index, style : Option[CellStyle] = style) : Cell =
    copy(value.asInstanceOf[Double], index, style)
}

case class BooleanCell private[model](value: Boolean, index: Option[Int], style: Option[CellStyle])
  extends Cell {
  def copyCell(value : Any = value, index : Option[Int] = index, style : Option[CellStyle] = style) : Cell =
    copy(value.asInstanceOf[Boolean], index, style)
}

case class DateCell private[model](value: Date, index: Option[Int], style: Option[CellStyle])
  extends Cell {
  def copyCell(value : Any = value, index : Option[Int] = index, style : Option[CellStyle] = style) : Cell =
    copy(value.asInstanceOf[Date], index, style)
}

case class CalendarCell private[model](value: Calendar, index: Option[Int], style: Option[CellStyle])
  extends Cell {
  def copyCell(value : Any = value, index : Option[Int] = index, style : Option[CellStyle] = style) : Cell =
    copy(value.asInstanceOf[Calendar], index, style)
}