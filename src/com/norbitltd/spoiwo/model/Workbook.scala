package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.model.enums.MissingCellPolicy

object Workbook extends Factory {

  private lazy val defaultActiveSheet =
    defaultPOIWorkbook.getActiveSheetIndex
  private lazy val defaultFirstVisibleTab =
    defaultPOIWorkbook.getFirstVisibleTab
  private lazy val defaultForceFormulaRecalculation =
    defaultPOIWorkbook.getForceFormulaRecalculation
  private lazy val defaultMissingCellPolicy =
    MissingCellPolicy.Undefined


  private lazy val defaultHidden = false
  private lazy val defaultSelectedTab = 0

  val Empty = apply()

  def apply(activeSheet: Int = defaultActiveSheet,
            firstVisibleTab: Int = defaultFirstVisibleTab,
            forceFormulaRecalculation: Boolean = defaultForceFormulaRecalculation,
            hidden: Boolean = defaultHidden,
            missingCellPolicy: MissingCellPolicy = defaultMissingCellPolicy,
            selectedTab: Int = defaultSelectedTab,
            sheets: Iterable[Sheet] = Nil): Workbook =
    Workbook(
      wrap(activeSheet, defaultActiveSheet),
      wrap(firstVisibleTab, defaultFirstVisibleTab),
      wrap(forceFormulaRecalculation, defaultForceFormulaRecalculation),
      wrap(hidden, defaultHidden),
      wrap(missingCellPolicy, defaultMissingCellPolicy),
      wrap(selectedTab, defaultSelectedTab),
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
