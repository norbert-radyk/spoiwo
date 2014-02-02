package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFSheet

object SheetProperties extends Factory {

  private lazy val defaultActiveCell = defaultPOISheet.getActiveCell
  private lazy val defaultAutoBreaks = defaultPOISheet.getAutobreaks
  private lazy val defaultDefaultColumnWidth = defaultPOISheet.getDefaultColumnWidth
  private lazy val defaultDefaultRowHeight = defaultPOISheet.getDefaultRowHeight

  private lazy val defaultDisplayFormulas = true
  private lazy val defaultDisplayGridLines = true
  private lazy val defaultDisplayGuts = defaultPOISheet.getDisplayGuts
  private lazy val defaultDisplayRowColHeadings = true
  private lazy val defaultDisplayZeros = true

  def apply(activeCell: String = defaultActiveCell,
            autoBreaks: Boolean = defaultAutoBreaks,
            defaultColumnWidth: Int = defaultDefaultColumnWidth,
            defaultRowHeight: Short = defaultDefaultRowHeight,
            displayFormulas: Boolean = defaultDisplayFormulas,
            displayGridLines: Boolean = defaultDisplayGridLines,
            displayGuts: Boolean = defaultDisplayGuts,
            displayRowColHeadings: Boolean = defaultDisplayRowColHeadings,
            displayZeros: Boolean = defaultDisplayZeros
             ): SheetProperties =
    SheetProperties(
      wrap(activeCell, defaultActiveCell),
      wrap(autoBreaks, defaultAutoBreaks),
      wrap(defaultColumnWidth, defaultDefaultColumnWidth),
      wrap(defaultRowHeight, defaultDefaultRowHeight),
      wrap(displayFormulas, defaultDisplayFormulas),
      wrap(displayGridLines, defaultDisplayGridLines),
      wrap(displayGuts, defaultDisplayGuts),
      wrap(displayRowColHeadings, defaultDisplayRowColHeadings),
      wrap(displayZeros, defaultDisplayZeros)
    )

}

case class SheetProperties private[excel](
                                           activeCell: Option[String],
                                           autoBreaks: Option[Boolean],
                                           defaultColumnWidth: Option[Int],
                                           defaultRowHeight: Option[Short],
                                           displayFormulas: Option[Boolean],
                                           displayGridLines: Option[Boolean],
                                           displayGuts: Option[Boolean],
                                           displayRowColHeadings: Option[Boolean],
                                           displayZeros: Option[Boolean]) {

  def withActiveCell(activeCell: String) =
    copy(activeCell = Option(activeCell))

  def withAutoBreaks(autoBreaks: Boolean) =
    copy(autoBreaks = Option(autoBreaks))

  def withDefaultColumnWidth(defaultColumnWidth: Int) =
    copy(defaultColumnWidth = Option(defaultColumnWidth))

  def withDefaultRowHeight(defaultRowHeight: Short) =
    copy(defaultRowHeight = Option(defaultRowHeight))

  def withDisplayFormulas(displayFormulas : Boolean) =
    copy(displayFormulas = Option(displayFormulas))

  def withDisplayGridLines(displayGridLines : Boolean) =
    copy(displayGridLines = Option(displayGridLines))

  def withDisplayGuts(displayGuts: Boolean) =
    copy(displayGuts = Option(displayGuts))

  def withDisplayRowColHeadings(displayRowColHeadings : Boolean) =
    copy(displayRowColHeadings = Option(displayRowColHeadings))

  def withDisplayZeros(displayZeros : Boolean) =
    copy(displayZeros = Option(displayZeros))


  def applyTo(sheet: XSSFSheet) {
    activeCell.foreach(sheet.setActiveCell)
    autoBreaks.foreach(sheet.setAutobreaks)
    defaultColumnWidth.foreach(sheet.setDefaultColumnWidth)
    defaultRowHeight.foreach(sheet.setDefaultRowHeight)
    displayFormulas.foreach(sheet.setDisplayFormulas)
    displayGridLines.foreach(sheet.setDisplayGridlines)
    displayGuts.foreach(sheet.setDisplayGuts)
    displayRowColHeadings.foreach(sheet.setDisplayRowColHeadings)
    displayZeros.foreach(sheet.setDisplayZeros)
  }

}
