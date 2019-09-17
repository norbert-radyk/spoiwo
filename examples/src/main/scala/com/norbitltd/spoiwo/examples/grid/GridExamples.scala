package com.norbitltd.spoiwo.examples.grid

import com.norbitltd.spoiwo.model.{Cell, CellRange, CellStyle, Color, Row, Sheet}
import com.norbitltd.spoiwo.model.enums.CellFill
import com.norbitltd.spoiwo.model.grid.Grid

object GridExamples {

  import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._

  /*
    These two are just helpers.
    The upper one creates an NxM grid with colored background.
    The second creates an NxM grid where the last row use a formula to copy the first row (and with colored background).
   */
  def generateGrid(n: Int, m: Int, c: Color): Grid = {
    val cell = Cell("").withStyle(CellStyle(fillForegroundColor = c, fillPattern = CellFill.Solid))
    val row  = Row().withCells(Seq.fill(m)(cell))
    Grid { _ =>
      Seq.fill(n)(row)
    }
  }

  def generateGridWithFormula(n: Int, m: Int, c: Color): Grid = {
    val cell = Cell("").withStyle(CellStyle(fillForegroundColor = c, fillPattern = CellFill.Solid))
    val row  = Row().withCells(Seq.fill(m)(cell))

    Grid { f =>
      def cellWithFormula(x: Int) =
        Cell(s"=${f.getCellByIndexes(x, 0)}") // <- this is the "tricky" part with the functions!
          .withStyle(CellStyle(fillForegroundColor = c, fillPattern = CellFill.Solid))
      val rowWithFormula = Row().withCells((1 to m).map(cellWithFormula))
      Seq.fill(n - 1)(row) ++ Seq(rowWithFormula)
    }
  }

  /*
    Actual examples!
   */
  def renderToRows(): Unit = {
    val rows = Grid.render(generateGrid(5, 5, Color.LightBlue))
    Sheet(name = "new sheet").withRows(rows).saveAsXlsx("test.xlsx")
  }
  def renderToRowsAndMergedCells(): Unit = {
    val grid = generateGrid(5, 5, Color.LightBlue)
    val gridWithMergedCells = Grid.addMergedRegion(grid, CellRange(2 -> 3, 2 -> 3))
    val rows = Grid.render(gridWithMergedCells)
    Sheet(name = "new sheet").withRows(rows).withMergedRegions(gridWithMergedCells.mergedRegions).saveAsXlsx("test.xlsx")
  }
  def theSameAsAboveButWithHelpers(): Unit = {
    import Grid._
    Sheet(name = "new sheet")
      .withGrid(generateGrid(5, 5, Color.LightBlue).addMergedRegion(CellRange(2 -> 3, 2 -> 3)))
      .saveAsXlsx("test.xlsx")
  }

  //This will create an 5x5 blue, a 2x2 green and a 7x7 read area besides each other.
  def horizontalCompose(): Unit = {
    import Grid._
    val g1 = generateGrid(5, 5, Color.LightBlue)
    val g2 = generateGrid(2, 2, Color.Green)
    val g3 = generateGrid(7, 7, Color.IndianRed)

    val composed = g1.addToRight(g3.addToLeft(g2))
    Sheet(name = "new sheet")
      .withGrid(composed)
      .saveAsXlsx("test.xlsx")
  }

  //This will create an 5x5 blue, a 2x2 green and a 7x7 read area under each other.
  def verticalCompose(): Unit = {
    import Grid._
    val g1 = generateGrid(5, 5, Color.LightBlue)
    val g2 = generateGrid(2, 2, Color.Green)
    val g3 = generateGrid(7, 7, Color.IndianRed)

    val composed = g1.addBelow(g3.addAbove(g2))
    Sheet(name = "new sheet")
      .withGrid(composed)
      .saveAsXlsx("test.xlsx")
  }

  //blue and green next to each other, the red below those, and the yellow next to all of the previous composition
  def biggerCompose(): Unit = {
    import Grid._
    val g1 = generateGrid(5, 5, Color.LightBlue)
    val g2 = generateGrid(2, 2, Color.Green)
    val g3 = generateGrid(7, 7, Color.IndianRed)
    val g4 = generateGrid(12, 8, Color.LightYellow)

    val composed  = g1.addToLeft(g2).addBelow(g3).addToLeft(g4)
    Sheet(name = "new sheet")
      .withGrid(composed)
      .saveAsXlsx("test.xlsx")
  }

  //if you write in the first line of the green areas the last line will copy the values
  def ifYouHaveFormulasThoseHandledToo(): Unit = {
    import Grid._
    val g1   = generateGrid(5, 5, Color.LightBlue)
    val g2   = generateGridWithFormula(6, 6, Color.Green)
    val composed  = g1.addToRight(g2).addBelow(g2)
    Sheet(name = "new sheet")
      .withGrid(composed)
      .saveAsXlsx("test.xlsx")
  }

  def main(args: Array[String]): Unit = {
    renderToRows()
    renderToRowsAndMergedCells()
    theSameAsAboveButWithHelpers()
    horizontalCompose()
    verticalCompose()
    biggerCompose()
    ifYouHaveFormulasThoseHandledToo()
  }
}
