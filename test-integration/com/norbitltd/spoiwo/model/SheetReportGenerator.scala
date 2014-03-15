package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.natures.csv.Model2CsvConversions._
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._

abstract class SheetReportGenerator {

  def getSheet : Sheet

  def main(args : Array[String]) {
    try {
      val xlsxFileName = args.head
      val csvFileName = args.tail.head
      println("Generating test report for XLSX:" + xlsxFileName + " and CSV: " + csvFileName)
      val sheet = getSheet
      sheet.saveAsXlsx(xlsxFileName)
      sheet.saveAsCsv(csvFileName)
      println("Report generated successfully!")
    } catch {
      case e : Exception => {
        println("Error generating test report, cause: " + e.getMessage, e)
        e.printStackTrace()
      }
    }

  }
}
