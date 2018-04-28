package com.norbitltd.spoiwo.model.enums

object PageOrder {

  lazy val DownThenOver = PageOrder("DownThenOver")
  lazy val OverThenDown = PageOrder("OverThenDown")
}

case class PageOrder(value: String) {

  override def toString: String = value

}
