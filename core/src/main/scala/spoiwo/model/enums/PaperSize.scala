package spoiwo.model.enums

object PaperSize {
  val Letter: PaperSize = PaperSize("Letter")
  val LetterSmall: PaperSize = PaperSize("LetterSmall")
  val Tabloid: PaperSize = PaperSize("Tabloid")
  val Ledger: PaperSize = PaperSize("Ledger")
  val Legal: PaperSize = PaperSize("Legal")
  val Statement: PaperSize = PaperSize("Statement")
  val Executive: PaperSize = PaperSize("Executive")
  val A3: PaperSize = PaperSize("A3")
  val A4: PaperSize = PaperSize("A4")
  val A4Small: PaperSize = PaperSize("A4Small")
  val A5: PaperSize = PaperSize("A5")
  val B4: PaperSize = PaperSize("B4")
  val B5: PaperSize = PaperSize("B5")
  val Folio: PaperSize = PaperSize("Folio")
  val Quarto: PaperSize = PaperSize("Quarto")
  val Standard10x14: PaperSize = PaperSize("Standard10x14")
  val Standard11x17: PaperSize = PaperSize("Standard11x17")

}

case class PaperSize private (value: String) {

  override def toString: String = value

}
