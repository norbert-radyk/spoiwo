package com.norbitltd.spoiwo.natures.xlsx

import scala.collection.JavaConverters._
import org.scalatest.FlatSpec
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import com.norbitltd.spoiwo.model._
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions.{convertTable, convertSheet}

class Model2XlsxConversionsForTableSpec extends FlatSpec {

  private def defaultTableWithoutId =
    Table(cellRange = CellRange(0 -> 1, 0 -> 1))

  private def defaultTable = defaultTableWithoutId.withId(1)
  private def defaultBook  = new XSSFWorkbook()
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
    val xlsx  = convertTable(model, defaultSheet)
    assert(xlsx.getCTTable.getId === specifiedId)
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
    val xlsx  = convertTable(model, defaultSheet)
    assert(xlsx.getDisplayName === specifiedDisplayName)
  }

  it should "have a generated display name when not specified" in {
    val tableId = 42
    val model = defaultTable.withId(42)
    val xlsx  = convertTable(model, defaultSheet)
    assert(xlsx.getDisplayName === s"Table$tableId")
  }

  it should "have the table name as specified" in {
    val specifiedName = "Bar"
    val model = defaultTable.withName(specifiedName)
    val xlsx  = convertTable(model, defaultSheet)
    assert(xlsx.getName === specifiedName)
  }

  it should "have a generated name when not specified" in {
    val tableId = 42
    val model = defaultTable.withId(42)
    val xlsx  = convertTable(model, defaultSheet)
    assert(xlsx.getName === s"ct_table_$tableId")
  }

  it should "have the table reference as specified" in {
    val model = defaultTable.withCellRange(CellRange(0 -> 2, 0 -> 3))
    val xlsx  = convertTable(model, defaultSheet)
    assert(xlsx.getCTTable.getRef === "A1:D3")
  }

  it should "have the columns when specified" in {
    val model   = defaultTable.withColumns(List(
      TableColumn("ColA", 42),
      TableColumn("ColB", 43)
    ))
    val xlsx        = convertTable(model, defaultSheet)
    val xlsxColumns = xlsx.getCTTable.getTableColumns

    model.columns.zipWithIndex.foreach {
      case (modelColumn, index) ⇒
        val xlsxCol = xlsxColumns.getTableColumnArray(index)
        assert(xlsxCol.getId   === modelColumn.id)
        assert(xlsxCol.getName === modelColumn.name)
    }
  }

  it should "have generated columns when not specified" in {
    val model       = defaultTable
    val xlsx        = convertTable(model, defaultSheet)
    val xlsxColumns = xlsx.getCTTable.getTableColumns

    assert(xlsxColumns.getCount === 2)
    (0 to 1).foreach { index ⇒
      val xlsxCol = xlsxColumns.getTableColumnArray(index)
      val expectedId = index + 1
      val expectedName = s"TableColumn$expectedId"
      assert(xlsxCol.getId   === expectedId)
      assert(xlsxCol.getName === expectedName)
    }
  }

  it should "have table style name as specified" in {
    val model     = defaultTable.withStyle(TableStyle(TableStyleName.TableStyleMedium2))
    val xlsx      = convertTable(model, defaultSheet)
    val xlsxStyle = xlsx.getCTTable.getTableStyleInfo

    assert(xlsxStyle.getName === TableStyleName.TableStyleMedium2.value)
    assert(xlsxStyle.getShowColumnStripes === false)
    assert(xlsxStyle.getShowRowStripes    === false)
  }

  it should "have table style with showRowStripes and showColumnStripes when specified" in {
    val model     = defaultTable.withStyle(
      TableStyle(TableStyleName.TableStyleMedium2, showColumnStripes = true, showRowStripes = true)
    )
    val xlsx      = convertTable(model, defaultSheet)
    val xlsxStyle = xlsx.getCTTable.getTableStyleInfo

    assert(xlsxStyle.getName === TableStyleName.TableStyleMedium2.value)
    assert(xlsxStyle.getShowColumnStripes === true)
    assert(xlsxStyle.getShowRowStripes    === true)
  }

  it should "have no table style when not specified" in {
    val model     = defaultTable
    val xlsx      = convertTable(model, defaultSheet)
    val xlsxStyle = xlsx.getCTTable.getTableStyleInfo

    assert(xlsxStyle === null)
  }

  it should "have autoFilter enabled when specified" in {
    val model          = defaultTable.withAutoFilter
    val xlsx           = convertTable(model, defaultSheet)
    val xlsxAutoFilter = xlsx.getCTTable.getAutoFilter

    assert(xlsxAutoFilter !== null)
  }

  it should "not have autoFilter enabled when not specified" in {
    val model          = defaultTable.withoutAutoFilter
    val xlsx           = convertTable(model, defaultSheet)
    val xlsxAutoFilter = xlsx.getCTTable.getAutoFilter

    assert(xlsxAutoFilter === null)
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
    val xlsx         = convertSheet(sheet, defaultBook)
    val xlsxTableIds = xlsx.getTables.asScala.map(_.getCTTable.getId)

    assert(xlsxTableIds.toSet === Set(1,2))
  }
}
