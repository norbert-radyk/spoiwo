package com.norbitltd.spoiwo.model

object CellDataFormat {

  val Undefined = apply(None)

  def apply(formatString: String): CellDataFormat = CellDataFormat(Option(formatString))

}

case class CellDataFormat private(formatString: Option[String])
