package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFSheet

object Column {

  val Default = Column()

}

case class Column(width: Int = 20, hidden: Boolean = false, groupCollapsed: Boolean = false, columnStyle: CellStyle = CellStyle.Default) {

  def withWidth(width: Int) = copy(width = width)

  def withHidden(hidden: Boolean) = copy(hidden = hidden)

  def withGroupCollapsed(groupCollapsed: Boolean) = copy(groupCollapsed = groupCollapsed)

  def withColumnStyle(columnStyle: CellStyle) = copy(columnStyle = columnStyle)

  def applyTo(index : Short, sheet: XSSFSheet) {
    sheet.setColumnWidth(index, 256 * width)
    sheet.setColumnHidden(index, hidden)
    sheet.setColumnGroupCollapsed(index, groupCollapsed)
    sheet.setDefaultColumnStyle(index, columnStyle.convert(sheet))
  }
}
