package com.norbitltd.spoiwo.model.enums

object Pane {

  lazy val LowerRightPane: Pane = Pane("LowerRightPane")
  lazy val LowerLeftPane: Pane = Pane("LowerLeftPane")
  lazy val UpperRightPane: Pane = Pane("UpperRightPane")
  lazy val UpperLeftPane: Pane = Pane("UpperLeftPane")

}

case class Pane private (value: String) {
  override def toString: String = value
}
