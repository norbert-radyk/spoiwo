package com.norbitltd.spoiwo.model.grid

import com.norbitltd.spoiwo.model.{CellRange, Row, Sheet, grid}

case class Grid(rows: FunctionHelper => Seq[Row], mergedRegions: Seq[CellRange] = Seq.empty) {
  lazy val height: Int = {
    rows(FunctionHelper.empty).foldLeft(0) { case (height, row) => row.index.map(_ + 1).getOrElse(height + 1) }
  }
  lazy val width: Int = {
    rows(FunctionHelper.empty)
      .map(_.cells.foldLeft(0) { case (width, cell) => cell.index.map(_ + 1).getOrElse(width + 1) })
      .max
  }
}

object Grid {
  import RowHelper._

  def addHorizontal(g1: Grid, g2: Grid): Grid = {
    val newFunc = (f: FunctionHelper) => {
      (g1.rows(f).fixupIndexes ++ g2.rows(f.moveRight(g1.width)).moveTo(g1.width, 0)).merge
        .sortBy(_.index.getOrElse(0))
    }
    val newMergedRegions = g1.mergedRegions ++ g2.mergedRegions.moveTo(g1.width, 0)
    grid.Grid(newFunc, newMergedRegions)
  }

  def addVertical(g1: Grid, g2: Grid): Grid = {
    val newFunc = (f: FunctionHelper) => {
      (g1.rows(f).fixupIndexes ++ g2.rows(f.moveDown(g1.height)).moveTo(0, g1.height)).merge
        .sortBy(_.index.getOrElse(0))
    }
    val newMergedRegions = g1.mergedRegions ++ g2.mergedRegions.moveTo(0, g1.height)
    grid.Grid(newFunc, newMergedRegions)
  }

  def render(g: Grid): Seq[Row] = {
    g.rows(new FunctionHelper(0, 0))
  }

  def addMergedRegion(g: Grid, mergedRegion: CellRange): Grid = {
    g.copy(g.rows, mergedRegion +: g.mergedRegions)
  }

  implicit class MergedRegionsHelper(mergedRegions: Seq[CellRange]) {
    def moveTo(x: Int, y: Int): Seq[CellRange] = {
      mergedRegions.map(
        r => CellRange((r.rowRange._1 + y, r.rowRange._2 + y), (r.columnRange._1 + x, r.columnRange._2 + x))
      )
    }
  }

  implicit class SheetExtender(val s: Sheet) extends AnyVal {
    def withGrid(g: Grid): Sheet = {
      s.withRows(render(g)).withMergedRegions(g.mergedRegions)
    }
  }

  implicit class GridBuilder(val g: Grid) extends AnyVal {
    def addToRight(g2: Grid): Grid = addHorizontal(g, g2)
    def addToLeft(g2: Grid): Grid = addHorizontal(g2, g)
    def addAbove(g2: Grid): Grid = addVertical(g2, g)
    def addBelow(g2: Grid): Grid = addVertical(g, g2)
    def render: Seq[Row] = Grid.render(g)
    def addMergedRegion(region: CellRange): Grid = Grid.addMergedRegion(g, region)
  }


}
