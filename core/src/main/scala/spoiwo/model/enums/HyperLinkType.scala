package spoiwo.model.enums

object HyperLinkType {
  lazy val Url: HyperLinkType = HyperLinkType("Url")
  lazy val Document: HyperLinkType = HyperLinkType("Document")
  lazy val Email: HyperLinkType = HyperLinkType("Email")
  lazy val File: HyperLinkType = HyperLinkType("File")
}

case class HyperLinkType private (value: String) {

  override def toString: String = value

}
