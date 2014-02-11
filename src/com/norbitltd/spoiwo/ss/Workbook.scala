package com.norbitltd.spoiwo.ss

import org.apache.poi.ss.usermodel.Row.MissingCellPolicy
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream
import scala.io.Source
import com.norbitltd.spoiwo.csv.CSVProperties
import com.norbitltd.spoiwo.utils.FileUtils

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

case class Workbook private[ss](
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

  /**
   * Converts the defined workbook into the sheet name -> csv content map for all of the sheets.
   * @return A sheet name -> CSV content map for each of the sheets
   */
  def convertToCSV(properties : CSVProperties = CSVProperties.Default) : Map[String, String] = {
    require(sheets.size <= 1 || sheets.forall(_.name.isDefined),
      "When converting workbook with multiple sheets to CSV format it is required to specify the unique name for each of them!")
    sheets.map(s => s.convertToCSV(properties)).toMap
  }

  def convertToXLSX(): XSSFWorkbook = {
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
  
  def saveAsXlsx(fileName: String) {
    val stream = new FileOutputStream(fileName)
    try {
      val workbook = convertToXLSX()
      workbook.write(stream)
    } finally {
      stream.flush()
      stream.close()
    }
  }

  def saveAsCsv(fileName : String, properties : CSVProperties = CSVProperties.Default) {
    val convertedCsvData = convertToCSV(properties)
    if( sheets.size <= 1 ) {
      convertedCsvData.values.foreach(csvContent => FileUtils.write(fileName, csvContent))
    } else {
      val fileNameCore = fileName.replace(".csv", "").replace(".CSV", "")
      convertedCsvData.foreach { case( sheetName, csvContent) =>
        val sheetFileName = fileNameCore + "_" + sheetName + ".csv"
        FileUtils.write(sheetFileName, csvContent)
      }
    }
  }

}
