package com.norbitltd.spoiwo.natures.xlsx

import com.norbitltd.spoiwo.model.HasIndex._
import com.norbitltd.spoiwo.model.{Cell, Row, Sheet}
import org.apache.poi.ss.usermodel

import scala.reflect.ClassTag

object Utils {
  def generateSheet[R, C](rowRange: Seq[R],
                          columnRange: Seq[C],
                          dataGenerator: (R, C) => String = (r: R, c: C) => s"$c,$r")(
      implicit rNum: Numeric[R],
      cNum: Numeric[C]
  ): Sheet = {
    val rows = rowRange.map { rowNum =>
      val cells = columnRange.map { colNum => Cell(dataGenerator(rowNum, colNum), index = cNum.toInt(colNum))
      }
      Row(cells, index = rNum.toInt(rowNum))
    }
    Sheet(rows: _*)
  }

  def mergeSheetData[T : ClassTag](sheets: Seq[Sheet], f: Cell => T): Array[Array[T]] = {
    val maxRowIndex = sheets.foldLeft(0)(_ max _.rows.maxIndex)
    val maxColIndex = sheets.foldLeft(0)(_ max _.rows.foldLeft(0)(_ max _.cells.maxIndex))
    val dataMatrix: Array[Array[T]] = Array.ofDim[T](maxRowIndex + 1, maxColIndex + 1)
    sheets.foreach(_.rows.foreach(r => r.cells.foreach(c => dataMatrix(r.index.get)(c.index.get) = f(c))))
    dataMatrix
  }

  def nonEqualCells[T](sheet: usermodel.Sheet, data: Array[Array[T]], extractor: usermodel.Cell => T)(implicit ev: Null <:< T): Seq[(T, T)] = {
    data.zipWithIndex.flatMap {
      case (dataRow, rowIdx) =>
        dataRow.zipWithIndex.flatMap {
          case (expectedData, cellIdx) =>
            val actualData = for {
              row <- Option(sheet.getRow(rowIdx))
              cell <- Option(row.getCell(cellIdx))
            } yield extractor(cell)
            if (actualData.orNull[T] == expectedData) Nil
            else List((actualData.orNull[T], expectedData))
        }
    }.toSeq
  }
}
