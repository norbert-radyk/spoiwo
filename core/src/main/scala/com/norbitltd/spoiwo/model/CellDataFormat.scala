package com.norbitltd.spoiwo.model

object CellDataFormat {

  def apply(formatString: String): CellDataFormat = CellDataFormat(Option(formatString))

}

case class CellDataFormat private (formatString: Option[String])
