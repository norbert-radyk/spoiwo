package com.norbitltd.spoiwo.natures.html

import com.norbitltd.spoiwo.model.{Cell, Row, Sheet}
import scala.xml.Elem

object Model2HtmlConversions {

  implicit class HtmlSheet(sheet : Sheet) {
    def asHtmlTableElem()  =
      <table>
        {sheet.rows.map(asHtmlRowElem)}
      </table>
  }

  //implicit def addHtmlSheetMethods(sheet: Sheet) = new HtmlSheet(sheet)

  implicit def asHtmlRowElem(row: Row) : Elem  =
    <tr>
      { row.cells.map(asHtmlCellElem) }
    </tr>

  implicit def asHtmlCellElem(cell: Cell) : Elem  =
    <td></td>

}
