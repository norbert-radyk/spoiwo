package com.norbitltd.spoiwo.natures.xlsx

import java.lang.reflect.Field

import com.norbitltd.spoiwo.model.Height._
import com.norbitltd.spoiwo.model._
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.scalatest.{FlatSpec, Matchers}
import scala.language.postfixOps

class Model2XlsxConversionsForWorkbookCacheSpec extends FlatSpec with Matchers {

    private val style = CellStyle(
        font = Font(height = 24 points, fontName = "Courier New", italic = true, strikeout = true),
        dataFormat = CellDataFormat("0.0")
    )

    private val victim = Workbook(
        Sheet(name = "Test Sheet", rows = List(Row(Seq(Cell("1", style = style), Cell("2", style = style)))))
    )

    private val conversions = Model2XlsxConversions

    private val cellStyleCacheField: Field = conversions.getClass.getDeclaredField("cellStyleCache")
    cellStyleCacheField.setAccessible(true)
    private val dataFormatCacheField: Field = conversions.getClass.getDeclaredField("dataFormatCache")
    dataFormatCacheField.setAccessible(true)
    private val fontCacheField: Field = conversions.getClass.getDeclaredField("fontCache")
    fontCacheField.setAccessible(true)

    "Cached workbook " should "be removed from cache after conversion" in {
        assertNotCached(victim.convertAsXlsx())
    }

    private def assertNotCached(workbook: XSSFWorkbook) {
        def getValue(field: Field) = field.get(conversions).asInstanceOf[collection.mutable.Map[XSSFWorkbook, _]]
        getValue(cellStyleCacheField).keySet should not contain workbook
        getValue(dataFormatCacheField).keySet should not contain workbook
        getValue(fontCacheField).keySet should not contain workbook
    }

}
