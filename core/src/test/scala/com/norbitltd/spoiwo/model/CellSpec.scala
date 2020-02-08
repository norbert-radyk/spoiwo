package com.norbitltd.spoiwo.model

import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class CellSpec extends AnyFlatSpec with Matchers {

  "Cell" should "be created as formula cell when initialized with string starting with '='" in {
    Cell("=3+2") shouldBe a[FormulaCell]
  }

  it should "be created as a string cell normal text content" in {
    Cell("Normal text") shouldBe a[StringCell]
  }

}
