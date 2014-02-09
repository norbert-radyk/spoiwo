package com.norbitltd.spoiwo.ss

import org.apache.poi.xssf.usermodel.XSSFDataFormat
import java.util.{Calendar, Date}

class T003Test {

  def creatingDateCells() {
    val dateCellStyle = CellStyle(dataFormat = CellDataFormat("m/d/yy h:mm"))
    Sheet(
      name = "new sheet",
      rows = Row(
        Cell(new Date()),
        Cell(new Date(), dateCellStyle),
        Cell(Calendar.getInstance(), dateCellStyle)
      ) :: Nil
    ).save("workbook.xlsx")
  }


  case class Book(author : String, price : Double)

  def printBooks(books : List[Book]) {
    val bookRows = books.map(book => Row(Cell(book.author), Cell(book.price)))
    Sheet(
      name = "Koki books",
      rows = Row(Cell("Autor"), Cell("Cena")) :: Row.Empty :: bookRows
    ).save("books.xlsx")
  }

}
