package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * Created by Norbert on 31/01/14.
 */
object Tester {

  def main( args : Array[String]) {
    val workbook = new XSSFWorkbook()
   /*
    private val DefaultActiveSheet = 0
  private val DefaultFirstVisibleTab = 0
  private val DefaultForceFormulaRecalculation = false
  private val DefaultHidden = false
  private val DefaultMissingCellPolicy = null
  private val DefaultSelectedTab = 0
    */

    println("Result: " + workbook.isHidden)
  }

}
