package com.norbitltd.spoiwo.natures.streaming.xlsx

import java.io.ByteArrayOutputStream

import com.norbitltd.spoiwo.model._
import com.norbitltd.spoiwo.natures.streaming.xlsx.Model2XlsxConversions._
import org.scalatest.{FlatSpec, Matchers}

class ExportWorkbookStreamingSpec extends FlatSpec with Matchers {

  "Writing workbook to an output stream" should "write non empty content" in {
    val baos = new ByteArrayOutputStream()
    Workbook(Sheet(name = "Test Sheet", rows = List.empty)).writeToOutputStream(baos)
    baos.size() should be > 0
  }

  "Writing workbook to an output stream" should "fail if tables are set" in {
    val baos = new ByteArrayOutputStream()
    intercept[IllegalStateException] {
      Workbook(Sheet(name = "Test Sheet", rows = List.empty, tables = List(Table(CellRange((1,1),(1,1)))))).writeToOutputStream(baos)
    }

    baos.size() should be (0)
  }


}
