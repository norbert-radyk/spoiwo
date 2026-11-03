package spoiwo.model

import org.scalatest.flatspec.AnyFlatSpec
import spoiwo.model.enums.CellStyleInheritance
import org.scalatest.matchers.should.Matchers

class CellSpec extends AnyFlatSpec with Matchers {

  "Cell" should "be created as formula cell when initialized with string starting with '='" in {
    Cell("=3+2") shouldBe a[FormulaCell]
  }

  it should "be created as a string cell normal text content" in {
    Cell("Normal text") shouldBe a[StringCell]
  }

  it should "be created as a string cell from text starting with '='" in {
    Cell.text("=HYPERLINK(\"https://example.com\")") shouldBe StringCell(
      "=HYPERLINK(\"https://example.com\")",
      None,
      None,
      CellStyleInheritance.CellThenRowThenColumnThenSheet
    )
  }

  it should "be created as a blank cell from null text" in {
    Cell.text(null) shouldBe a[BlankCell]
  }

  it should "be created as a formula cell with or without the leading '='" in {
    Cell.formula("=SUM(A1:A3)") shouldBe Cell.formula("SUM(A1:A3)")
    Cell.formula("SUM(A1:A3)") shouldBe a[FormulaCell]
  }

}
