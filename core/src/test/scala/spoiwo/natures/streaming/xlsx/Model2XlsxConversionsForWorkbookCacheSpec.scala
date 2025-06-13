package spoiwo.natures.streaming.xlsx

import spoiwo.model._

import java.lang.reflect.Field
import Model2XlsxConversions._
import org.apache.poi.xssf.streaming.SXSSFWorkbook

import scala.language.postfixOps
import scala.util.Try
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers
import spoiwo.model.Height.HeightEnrichment

class Model2XlsxConversionsForWorkbookCacheSpec extends AnyFlatSpec with Matchers {

  private val style = CellStyle(
    font = Font(height = 24 points, fontName = "Courier New", italic = true, strikeout = true),
    dataFormat = CellDataFormat("0.0")
  )

  private val victim = Workbook(
    Sheet(name = "Test Sheet", rows = List(Row(Seq(Cell("1", style = style), Cell("2", style = style)))))
  )

  private val conversions = Model2XlsxConversions

  private val cellStyleCacheField: Field = getConversionsField("cellStyleCache")
  cellStyleCacheField.setAccessible(true)
  private val dataFormatCacheField: Field = getConversionsField("dataFormatCache")
  dataFormatCacheField.setAccessible(true)
  private val fontCacheField: Field = getConversionsField("fontCache")
  fontCacheField.setAccessible(true)

  "Cached workbook " should "be removed from cache after conversion" in {
    assertNotCached(victim.convertAsXlsx())
  }

  //Workaround for the scala/2.11 compiler which messes up with the private variable names
  private def getConversionsField(name: String): Field = Try(conversions.getClass.getDeclaredField(name)).getOrElse {
    val allFields = conversions.getClass.getDeclaredFields
    allFields.filter(_.getName.contains(name)).head
  }

  private def assertNotCached(workbook: SXSSFWorkbook): Unit = {
    def getValue(field: Field) = field.get(conversions).asInstanceOf[collection.mutable.Map[SXSSFWorkbook, ?]]
    getValue(cellStyleCacheField).keySet should not contain workbook
    getValue(dataFormatCacheField).keySet should not contain workbook
    getValue(fontCacheField).keySet should not contain workbook
  }

}
