package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.model.enums.MissingCellPolicy

object Workbook {

  val Empty = apply()

  def apply(activeSheet: java.lang.Integer = null,
            firstVisibleTab: java.lang.Integer = null,
            forceFormulaRecalculation: java.lang.Boolean = null,
            hidden: java.lang.Boolean = null,
            missingCellPolicy: MissingCellPolicy = null,
            selectedTab: java.lang.Integer = null,
            sheets: Iterable[Sheet] = Nil): Workbook =
    Workbook(
      Option(activeSheet).map(_.intValue),
      Option(firstVisibleTab).map(_.intValue),
      Option(forceFormulaRecalculation).map(_.booleanValue),
      Option(hidden).map(_.booleanValue),
      Option(missingCellPolicy),
      Option(selectedTab).map(_.intValue),
      sheets
    )

  def apply(sheets: Sheet*): Workbook =
    Workbook(sheets = sheets)
}

case class Workbook private[model](
                                    activeSheet: Option[Int],
                                    firstVisibleTab: Option[Int],
                                    forceFormulaRecalculation: Option[Boolean],
                                    hidden: Option[Boolean],
                                    missingCellPolicy: Option[MissingCellPolicy],
                                    selectedTab: Option[Int],
                                    sheets: Iterable[Sheet]) {


  def withActiveSheet(activeSheet: Int) =
    copy(activeSheet = Option(activeSheet))

  def withFirstVisibleTab(firstVisibleTab: Int) =
    copy(firstVisibleTab = Option(firstVisibleTab))

  def withForceFormulaRecalculation(forceFormulaRecalculation: Boolean) =
    copy(forceFormulaRecalculation = Option(forceFormulaRecalculation))

  def withHidden(hidden: Boolean) =
    copy(hidden = Option(hidden))

  def withMissingCellPolicy(missingCellPolicy: MissingCellPolicy) =
    copy(missingCellPolicy = Option(missingCellPolicy))

  def withSelectedTab(selectedTab: Int) =
    copy(selectedTab = Option(selectedTab))

  def withSheets(sheets: Iterable[Sheet]) =
    copy(sheets = sheets)

  def withSheets(sheets: Sheet*) =
    copy(sheets = sheets)

}
