package com.norbitltd.spoiwo.model.enums

object PaperSize {
  val Letter = PaperSize("Letter")
  val LetterSmall = PaperSize("LetterSmall")
  val Tabloid = PaperSize("Tabloid")
  val Ledger = PaperSize("Ledger")
  val Legal = PaperSize("Legal")
  val Statement = PaperSize("Statement")
  val Executive = PaperSize("Executive")
  val A3 = PaperSize("A3")
  val A4 = PaperSize("A4")
  val A4Small = PaperSize("A4Small")
  val A5 = PaperSize("A5")
  val B4 = PaperSize("B4")
  val B5 = PaperSize("B5")
  val Folio = PaperSize("Folio")
  val Quarto = PaperSize("Quarto")
  val Standard10x14 = PaperSize("Standard10x14")
  val Standard11x17 = PaperSize("Standard11x17")

}

case class PaperSize private (value: String) {

  override def toString: String = value

}
