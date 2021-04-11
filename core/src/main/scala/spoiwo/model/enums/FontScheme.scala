package spoiwo.model.enums

object FontScheme {
  lazy val None: FontScheme = FontScheme("None")
  lazy val Major: FontScheme = FontScheme("Major")
  lazy val Minor: FontScheme = FontScheme("Minor")
}

case class FontScheme private (value: String) {

  override def toString: String = value

}
