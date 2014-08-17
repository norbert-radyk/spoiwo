package com.norbitltd.spoiwo.model

import org.scalatest.FlatSpec


class CellSpec extends FlatSpec {

  "Cell" should "be created as formula cell when initialized with string starting with '='" in {
    val c = Cell("=3+2")
    assert(c.isInstanceOf[FormulaCell])
  }

  it should "be created as a string cell normal text content" in {
    val c = Cell("Normal text")
    assert(c.isInstanceOf[StringCell])
  }

}
