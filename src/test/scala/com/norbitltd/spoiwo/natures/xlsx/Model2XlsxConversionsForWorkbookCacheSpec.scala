package com.norbitltd.spoiwo.natures.xlsx

import java.lang.reflect.Field

import com.norbitltd.spoiwo.model.Height._
import com.norbitltd.spoiwo.model._
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.scalatest.FlatSpec
import scala.language.postfixOps

class Model2XlsxConversionsForWorkbookCacheSpec extends FlatSpec {


    val style = CellStyle(font = Font(height = 24 points, fontName = "Courier New", italic = true, strikeout = true), dataFormat = CellDataFormat("0.0"))
    private val victim = Workbook(Sheet(name = "Test Cheet", rows = List(Row(Seq(Cell("1", style = style), Cell("2", style = style))))))
    val conversions = Model2XlsxConversions
    val cellStyleCacheField = conversions.getClass.getDeclaredField("cellStyleCache")
    cellStyleCacheField.setAccessible(true)

    val dataFormatCacheField = conversions.getClass.getDeclaredField("com$norbitltd$spoiwo$natures$xlsx$Model2XlsxConversions$$dataFormatCache")
    dataFormatCacheField.setAccessible(true)

    val fontCacheField = conversions.getClass.getDeclaredField("fontCache")
    fontCacheField.setAccessible(true)


    "Cached workbook " should "be removed from cache after conversion" in {
        assertNotCached(victim.convertAsXlsx())
    }

    private def assertNotCached(workbook: XSSFWorkbook) {
        def getValue(field: Field) = field.get(conversions).asInstanceOf[collection.mutable.Map[XSSFWorkbook, _]]
        assert(!getValue(cellStyleCacheField).keySet.contains(workbook))
        assert(!getValue(dataFormatCacheField).keySet.contains(workbook))
        assert(!getValue(fontCacheField).keySet.contains(workbook))
    }

}
