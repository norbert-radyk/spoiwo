package com.norbitltd.spoiwo.model.grid


case class FunctionHelper(x: Int, y: Int) {
  def getRelativeCol(relativeX: Int): String = {
    FunctionHelper.colToLetter(x + relativeX - 1)
  }
  def getRelativeRow(relativeY: Int): String = {
    (y + relativeY + 1).toString
  }
  def getCellByIndexes(x: Int, y: Int): String = {
    getRelativeCol(x) + getRelativeRow(y)
  }

  protected[grid] def moveRight(x2: Int) = FunctionHelper(x2 + x, y)
  protected[grid] def moveDown(y2: Int)  = FunctionHelper(x, y + y2)
}

object FunctionHelper {
  def empty = FunctionHelper(0, 0)

  private def colToLetter(x: Int): String = {
    val simple = x + 'A'
    if (simple <= 'Z') {
      simple.toChar.toString
    } else {
      val remainder = x % ('Z' - 'A')
      val first     = x / ('Z' - 'A')
      (first + 'A').toChar.toString + (remainder + 'A').toChar.toString
    }
  }
}
