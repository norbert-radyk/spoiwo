package com.norbitltd.spoiwo.model

import org.apache.commons.logging.LogFactory
import com.norbitltd.spoiwo.natures.csv.Model2CsvConversions._
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._

abstract class AbstractReportGenerator {

  def log = LogFactory.getLog(this.getClass)

  def getWorkbook : Workbook

  def main(args : Array[String]) {
    try {
      val xlsxFileName = args.head
      val csvFileName = args.tail.head
      println("Generating test report for XLSX:" + xlsxFileName + " and CSV: " + csvFileName)
      val workbook = getWorkbook
      workbook.saveAsXlsx(xlsxFileName)
      workbook.saveAsCsv(csvFileName)
      println("Report generated successfully!")
    } catch {
      case e : Exception => {
        println("Error generating test report, cause: " + e.getMessage, e)
        e.printStackTrace()
      }
    }

  }

}
