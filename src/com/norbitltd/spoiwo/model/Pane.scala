package com.norbitltd.spoiwo.model

object Pane {

  lazy val LowerRightPane = Pane(1)
  lazy val LowerLeftPane = Pane(2)
  lazy val UpperRightPane = Pane(3)
  lazy val UpperLeftPane = Pane(4)

}

case class Pane private(id : Int)





