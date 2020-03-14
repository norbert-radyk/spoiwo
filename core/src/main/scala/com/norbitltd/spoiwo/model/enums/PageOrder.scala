package com.norbitltd.spoiwo.model.enums

object PageOrder {

  lazy val DownThenOver: PageOrder = PageOrder("DownThenOver")
  lazy val OverThenDown: PageOrder = PageOrder("OverThenDown")
}

case class PageOrder(value: String) {

  override def toString: String = value

}
