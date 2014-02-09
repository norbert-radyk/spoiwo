package com.norbitltd.spoiwo.ss

import org.apache.commons.logging.LogFactory

abstract class AbstractReportGenerator {

  def log = LogFactory.getLog(this.getClass)

  def getWorkbook : Workbook

  def main(args : Array[String]) {
    try {
      val xlsxFileName = args.head
      val csvFileName = args.tail.head
      println("Generating test report for XLSX:" + xlsxFileName + " and CSV: " + csvFileName)
      val workbook = getWorkbook
      workbook.saveXLSX(xlsxFileName)
      workbook.saveCSV(csvFileName)
      println("Report generated successfully!")
    } catch {
      case e : Exception => {
        println("Error generating test report, cause: " + e.getMessage, e)
        e.printStackTrace()
      }
    }

  }

}
