package com.norbitltd.spoiwo.natures.xlsx

import scala.collection.JavaConverters._
import org.scalatest.{FlatSpec, Matchers}
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import com.norbitltd.spoiwo.model._
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions.{convertSheet, convertTable}

class Model2XlsxConversionsForTableSpec extends FlatSpec with Matchers {

  private def defaultTableWithoutId =
    Table(cellRange = CellRange(0 -> 1, 0 -> 1))

  private def defaultTable = defaultTableWithoutId.withId(1)

  private def defaultBook = new XSSFWorkbook()

  private def defaultSheet = defaultBook.createSheet()

  "Table conversions" should "validate that specified columns match with cell reference" in {
    val model = defaultTable.withColumns(List(TableColumn("col1", 1)))
    intercept[IllegalArgumentException] {
      convertTable(model, defaultSheet)
    }
  }

  it should "have the table id as specified" in {
    val specifiedId = 10L
    val model = defaultTable.withId(specifiedId)
    val xlsx = convertTable(model, defaultSheet)
    xlsx.getCTTable.getId shouldBe specifiedId
  }

  it should "throw IllegalArgumentException if id not specified" in {
    intercept[IllegalArgumentException] {
      val model = defaultTableWithoutId
      convertTable(model, defaultSheet)
    }
  }

  it should "have the table display name as specified" in {
    val specifiedDisplayName = "Foo"
    val model = defaultTable.withDisplayName(specifiedDisplayName)
    val xlsx = convertTable(model, defaultSheet)
    xlsx.getDisplayName shouldBe specifiedDisplayName
  }

  it should "have a generated display name when not specified" in {
    val tableId = 42
    val model = defaultTable.withId(42)
    val xlsx = convertTable(model, defaultSheet)
    xlsx.getDisplayName shouldBe s"Table$tableId"
  }

  it should "have the table name as specified" in {
    val specifiedName = "Bar"
    val model = defaultTable.withName(specifiedName)
    val xlsx = convertTable(model, defaultSheet)
    xlsx.getName shouldBe specifiedName
  }

  it should "have a generated name when not specified" in {
    val tableId = 42
    val model = defaultTable.withId(42)
    val xlsx = convertTable(model, defaultSheet)
    xlsx.getName shouldBe s"ct_table_$tableId"
  }

  it should "have the table reference as specified" in {
    val model = defaultTable.withCellRange(CellRange(0 -> 2, 0 -> 3))
    val xlsx = convertTable(model, defaultSheet)
    xlsx.getCTTable.getRef shouldBe "A1:D3"
  }

  it should "have the columns when specified" in {
    val model = defaultTable.withColumns(
      List(
        TableColumn("ColA", 42),
        TableColumn("ColB", 43)
      ))
    val xlsx = convertTable(model, defaultSheet)
    val xlsxColumns = xlsx.getCTTable.getTableColumns

    model.columns.zipWithIndex.foreach {
      case (modelColumn, index) =>
        val xlsxCol = xlsxColumns.getTableColumnArray(index)
        xlsxCol.getId shouldBe modelColumn.id
        xlsxCol.getName shouldBe modelColumn.name
    }
  }

  it should "have generated columns when not specified" in {
    val model = defaultTable
    val xlsx = convertTable(model, defaultSheet)
    val xlsxColumns = xlsx.getCTTable.getTableColumns

    xlsxColumns.getCount shouldBe 2
    (0 to 1).foreach { index =>
      val xlsxCol = xlsxColumns.getTableColumnArray(index)
      val expectedId = index + 1
      val expectedName = s"TableColumn$expectedId"
      xlsxCol.getId shouldBe expectedId
      xlsxCol.getName shouldBe expectedName
    }
  }

  it should "have table style name as specified" in {
    val model = defaultTable.withStyle(TableStyle(TableStyleName.TableStyleMedium2))
    val xlsx = convertTable(model, defaultSheet)
    val xlsxStyle = xlsx.getCTTable.getTableStyleInfo

    xlsxStyle.getName shouldBe TableStyleName.TableStyleMedium2.value
    xlsxStyle.getShowColumnStripes shouldBe false
    xlsxStyle.getShowRowStripes shouldBe false
  }

  it should "have table style with showRowStripes and showColumnStripes when specified" in {
    val model = defaultTable.withStyle(
      TableStyle(TableStyleName.TableStyleMedium2, showColumnStripes = true, showRowStripes = true)
    )
    val xlsx = convertTable(model, defaultSheet)
    val xlsxStyle = xlsx.getCTTable.getTableStyleInfo

    xlsxStyle.getName shouldBe TableStyleName.TableStyleMedium2.value
    xlsxStyle.getShowColumnStripes shouldBe true
    xlsxStyle.getShowRowStripes shouldBe true
  }

  it should "have no table style when not specified" in {
    val model = defaultTable
    val xlsx = convertTable(model, defaultSheet)
    val xlsxStyle = xlsx.getCTTable.getTableStyleInfo

    xlsxStyle shouldBe null
  }

  it should "have autoFilter enabled when specified" in {
    val model = defaultTable.withAutoFilter
    val xlsx = convertTable(model, defaultSheet)
    val xlsxAutoFilter = xlsx.getCTTable.getAutoFilter

    xlsxAutoFilter should not be null
  }

  it should "not have autoFilter enabled when not specified" in {
    val model = defaultTable.withoutAutoFilter
    val xlsx = convertTable(model, defaultSheet)
    val xlsxAutoFilter = xlsx.getCTTable.getAutoFilter

    xlsxAutoFilter shouldBe null
  }

  it should "validate that specified table ids are unique" in {
    val sheet = {
      val t1 = defaultTable.withId(1)
      val t2 = defaultTable.withId(1)
      Sheet().withTables(t1, t2)
    }

    intercept[IllegalArgumentException] {
      convertSheet(sheet, defaultBook)
    }
  }

  it should "have generated table ids when not specified" in {
    val sheet = {
      val t1 = defaultTableWithoutId
      val t2 = defaultTableWithoutId
      Sheet().withTables(t1, t2)
    }
    val xlsx = convertSheet(sheet, defaultBook)
    val xlsxTableIds = xlsx.getTables.asScala.map(_.getCTTable.getId)

    xlsxTableIds.toSet shouldBe Set(1, 2)
  }
}
