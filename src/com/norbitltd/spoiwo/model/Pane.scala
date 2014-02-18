package com.norbitltd.spoiwo.model

import org.apache.poi.xssf.usermodel.XSSFSheet

object Pane {

  object LowerRightPane extends Pane
  object LowerLeftPane extends Pane
  object UpperRightPane extends Pane
  object UpperLeftPane extends Pane

}

case class Pane private() {

  import Pane._

  def convert(): Int = this match {
    case LowerLeftPane => org.apache.poi.ss.usermodel.Sheet.PANE_LOWER_LEFT
    case LowerRightPane => org.apache.poi.ss.usermodel.Sheet.PANE_LOWER_RIGHT
    case UpperLeftPane => org.apache.poi.ss.usermodel.Sheet.PANE_UPPER_LEFT
    case UpperRightPane => org.apache.poi.ss.usermodel.Sheet.PANE_UPPER_RIGHT
  }
}

sealed trait PaneAction {

  def applyTo(sheet: XSSFSheet): Unit

}

case class NoSplitOrFreeze() extends PaneAction {
  override def applyTo(sheet: XSSFSheet) {
    sheet.createFreezePane(0, 0)
  }
}

object FreezePane {
  def apply(columnSplit: Int, rowSplit: Int) : FreezePane =
    FreezePane(columnSplit, rowSplit, columnSplit, rowSplit)
}

case class FreezePane(columnSplit: Int,
                      rowSplit: Int,
                      leftMostColumn: Int,
                      topRow: Int) extends PaneAction {

  override def applyTo(sheet: XSSFSheet) {
    sheet.createFreezePane(columnSplit, rowSplit, leftMostColumn, topRow)
  }

}

case class SplitPane(
                      xSplitPosition: Int,
                      ySplitPosition: Int,
                      leftMostColumn: Int = 0,
                      topRow: Int = 0,
                      activePane: Pane = Pane.UpperLeftPane) extends PaneAction {

  override def applyTo(sheet: XSSFSheet) {
    sheet.createSplitPane(xSplitPosition, ySplitPosition, leftMostColumn, topRow, activePane.convert())
  }

}
