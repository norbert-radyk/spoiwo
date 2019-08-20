package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.model.enums.Pane

sealed trait PaneAction

case class NoSplitOrFreeze() extends PaneAction

case class FreezePane(columnSplit: Int, rowSplit: Int, leftMostColumn: Int, topRow: Int) extends PaneAction

case class SplitPane(xSplitPosition: Int,
                     ySplitPosition: Int,
                     leftMostColumn: Int = 0,
                     topRow: Int = 0,
                     activePane: Pane = Pane.UpperLeftPane)
    extends PaneAction

object FreezePane {
  def apply(columnSplit: Int, rowSplit: Int): FreezePane =
    FreezePane(columnSplit, rowSplit, columnSplit, rowSplit)
}
