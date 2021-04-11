package spoiwo.model.enums

object MissingCellPolicy {
  lazy val ReturnNullAndBlank: MissingCellPolicy = MissingCellPolicy("ReturnNullAndBlank")
  lazy val ReturnBlankAsNull: MissingCellPolicy = MissingCellPolicy("ReturnBlankAsNull")
  lazy val CreateNullAsBlank: MissingCellPolicy = MissingCellPolicy("CreateNullAsBlank")
}

case class MissingCellPolicy private (value: String) {
  override def toString: String = value
}
