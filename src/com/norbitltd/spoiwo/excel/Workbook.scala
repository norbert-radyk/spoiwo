package com.norbitltd.spoiwo.excel

import org.apache.poi.ss.usermodel.Row.MissingCellPolicy
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream

object Workbook extends Factory {

  private lazy val defaultActiveSheet =
    defaultPOIWorkbook.getActiveSheetIndex
  private lazy val defaultFirstVisibleTab =
    defaultPOIWorkbook.getFirstVisibleTab
  private lazy val defaultForceFormulaRecalculation =
    defaultPOIWorkbook.getForceFormulaRecalculation
  private lazy val defaultMissingCellPolicy =
    defaultPOIWorkbook.getMissingCellPolicy

  //TODO defaultPOIWorkbook.isHidden throw Not implemented yet exception - to be replaced in the future.
  private lazy val defaultHidden = false
  //TODO no getSelectedTab method not available on Apache POI, so using sensible default.
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

case class Workbook private[excel](
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

  def convert(): XSSFWorkbook = {
    val workbook = new XSSFWorkbook()
    sheets.foreach(sheet => sheet.convert(workbook))

    //Parameters
    activeSheet.foreach(workbook.setActiveSheet)
    firstVisibleTab.foreach(workbook.setFirstVisibleTab)
    forceFormulaRecalculation.foreach(workbook.setForceFormulaRecalculation)
    hidden.foreach(workbook.setHidden)
    missingCellPolicy.foreach(workbook.setMissingCellPolicy)
    selectedTab.foreach(workbook.setSelectedTab)
    workbook
  }

  def save(fileName: String) = {
    val stream = new FileOutputStream(fileName)
    try {
      val workbook = convert()
      workbook.write(stream)
    } finally {
      stream.flush()
      stream.close()
    }
  }
}
