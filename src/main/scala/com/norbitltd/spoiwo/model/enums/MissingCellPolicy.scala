package com.norbitltd.spoiwo.model.enums

object MissingCellPolicy {
  lazy val ReturnNullAndBlank = MissingCellPolicy("ReturnNullAndBlank")
  lazy val ReturnBlankAsNull = MissingCellPolicy("ReturnBlankAsNull")
  lazy val CreateNullAsBlank = MissingCellPolicy("CreateNullAsBlank")
}

case class MissingCellPolicy private (value: String) {
  override def toString: String = value
}
