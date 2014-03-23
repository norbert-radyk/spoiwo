package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.model.enums.Pane

object Test {
  def main(args: Array[String]) {
    /*println(Color.Undefined == Color(0, 0, 0))
    println(Color.Undefined == Color.Black)
    println(Color.Undefined == Color.Undefined)
    println(Color(0,0,0) == Color.Black)*/
    println(CellBorderStyle.DashDotDot == CellBorderStyle.Dotted)
    println(Pane.LowerLeftPane == Pane.UpperLeftPane)
  }
}
