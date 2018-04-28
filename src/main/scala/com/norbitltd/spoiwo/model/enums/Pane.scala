package com.norbitltd.spoiwo.model.enums

object Pane {

  lazy val LowerRightPane = Pane("LowerRightPane")
  lazy val LowerLeftPane = Pane("LowerLeftPane")
  lazy val UpperRightPane = Pane("UpperRightPane")
  lazy val UpperLeftPane = Pane("UpperLeftPane")

}

case class Pane private (value: String) {
  override def toString: String = value
}
