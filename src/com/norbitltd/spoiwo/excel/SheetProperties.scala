package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.{XSSFSheet}

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
  private lazy val defaultFitToPage = defaultPOISheet.getFitToPage
  private lazy val defaultForceFormulaRecalculation = defaultPOISheet.getForceFormulaRecalculation
  private lazy val defaultHorizontallyCenter = defaultPOISheet.getHorizontallyCenter
  private lazy val defaultPrintGridLines = false
  private lazy val defaultRightToLeft = false
  private lazy val defaultRowSumsBelow = defaultPOISheet.getRowSumsBelow
  private lazy val defaultRowSumsRight = defaultPOISheet.getRowSumsRight
  private lazy val defaultSelected = false
  private lazy val defaultTabColor = Color.WHITE
  private lazy val defaultVirtuallyCenter = defaultPOISheet.getVerticallyCenter
  private lazy val defaultZoom = 100

  lazy val Default = SheetProperties()

  def apply(activeCell: String = defaultActiveCell,
            autoBreaks: Boolean = defaultAutoBreaks,
            defaultColumnWidth: Int = defaultDefaultColumnWidth,
            defaultRowHeight: Short = defaultDefaultRowHeight,
            displayFormulas: Boolean = defaultDisplayFormulas,
            displayGridLines: Boolean = defaultDisplayGridLines,
            displayGuts: Boolean = defaultDisplayGuts,
            displayRowColHeadings: Boolean = defaultDisplayRowColHeadings,
            displayZeros: Boolean = defaultDisplayZeros,
            fitToPage: Boolean = defaultFitToPage,
            forceFormulaRecalculation: Boolean = defaultForceFormulaRecalculation,
            horizontallyCenter: Boolean = defaultHorizontallyCenter,
            printGridLines: Boolean = defaultPrintGridLines,
            rightToLeft: Boolean = defaultRightToLeft,
            rowSumsBelow: Boolean = defaultRowSumsBelow,
            rowSumsRight: Boolean = defaultRowSumsRight,
            selected: Boolean = defaultSelected,
            tabColor: Color = defaultTabColor,
            virtuallyCenter: Boolean = defaultVirtuallyCenter,
            zoom: Int = defaultZoom

             ): SheetProperties = SheetProperties(
    wrap(activeCell, defaultActiveCell),
    wrap(autoBreaks, defaultAutoBreaks),
    wrap(defaultColumnWidth, defaultDefaultColumnWidth),
    wrap(defaultRowHeight, defaultDefaultRowHeight),
    wrap(displayFormulas, defaultDisplayFormulas),
    wrap(displayGridLines, defaultDisplayGridLines),
    wrap(displayGuts, defaultDisplayGuts),
    wrap(displayRowColHeadings, defaultDisplayRowColHeadings),
    wrap(displayZeros, defaultDisplayZeros),
    wrap(fitToPage, defaultFitToPage),
    wrap(forceFormulaRecalculation, defaultForceFormulaRecalculation),
    wrap(horizontallyCenter, defaultHorizontallyCenter),
    wrap(printGridLines, defaultPrintGridLines),
    wrap(rightToLeft, defaultRightToLeft),
    wrap(rowSumsBelow, defaultRowSumsBelow),
    wrap(rowSumsRight, defaultRowSumsRight),
    wrap(selected, defaultSelected),
    wrap(tabColor, defaultTabColor),
    wrap(virtuallyCenter, defaultVirtuallyCenter),
    wrap(zoom, defaultZoom)
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
                                           displayZeros: Option[Boolean],
                                           fitToPage: Option[Boolean],
                                           forceFormulaRecalculation: Option[Boolean],
                                           horizontallyCenter: Option[Boolean],
                                           printGridLines: Option[Boolean],
                                           rightToLeft: Option[Boolean],
                                           rowSumsBelow: Option[Boolean],
                                           rowSumsRight: Option[Boolean],
                                           selected: Option[Boolean],
                                           tabColor: Option[Color],
                                           virtuallyCenter: Option[Boolean],
                                           zoom: Option[Int]) {

  def withActiveCell(activeCell: String) =
    copy(activeCell = Option(activeCell))

  def withAutoBreaks(autoBreaks: Boolean) =
    copy(autoBreaks = Option(autoBreaks))

  def withDefaultColumnWidth(defaultColumnWidth: Int) =
    copy(defaultColumnWidth = Option(defaultColumnWidth))

  def withDefaultRowHeight(defaultRowHeight: Short) =
    copy(defaultRowHeight = Option(defaultRowHeight))

  def withDisplayFormulas(displayFormulas: Boolean) =
    copy(displayFormulas = Option(displayFormulas))

  def withDisplayGridLines(displayGridLines: Boolean) =
    copy(displayGridLines = Option(displayGridLines))

  def withDisplayGuts(displayGuts: Boolean) =
    copy(displayGuts = Option(displayGuts))

  def withDisplayRowColHeadings(displayRowColHeadings: Boolean) =
    copy(displayRowColHeadings = Option(displayRowColHeadings))

  def withDisplayZeros(displayZeros: Boolean) =
    copy(displayZeros = Option(displayZeros))

  def withFitToPage(fitToPage: Boolean) =
    copy(fitToPage = Option(fitToPage))

  def withForceFormulaRecalculation(forceFormulaRecalculation: Boolean) =
    copy(forceFormulaRecalculation = Option(forceFormulaRecalculation))

  def withHorizontallyCenter(horizontallyCenter: Boolean) =
    copy(horizontallyCenter = Option(horizontallyCenter))

  def withPrintGridLines(printGridLines: Boolean) =
    copy(printGridLines = Option(printGridLines))

  def withRightToLeft(rightToLeft: Boolean) =
    copy(rightToLeft = Option(rightToLeft))

  def withRowSumsBelow(rowSumsBelow: Boolean) =
    copy(rowSumsBelow = Option(rowSumsBelow))

  def withRowSumsRight(rowSumsRight: Boolean) =
    copy(rowSumsRight = Option(rowSumsRight))

  def withSelected(selected: Boolean) =
    copy(selected = Option(selected))

  def withTabColor(tabColor: Color) =
    copy(tabColor = Option(tabColor))

  def withVirtuallyCenter(virtuallyCenter: Boolean) =
    copy(virtuallyCenter = Option(virtuallyCenter))

  def withZoom(zoom: Int) =
    copy(zoom = Option(zoom))

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
    fitToPage.foreach(sheet.setFitToPage)
    forceFormulaRecalculation.foreach(sheet.setForceFormulaRecalculation)
    horizontallyCenter.foreach(sheet.setHorizontallyCenter)
    printGridLines.foreach(sheet.setPrintGridlines)
    rightToLeft.foreach(sheet.setRightToLeft)
    rowSumsBelow.foreach(sheet.setRowSumsBelow)
    rowSumsRight.foreach(sheet.setRowSumsRight)
    selected.foreach(sheet.setSelected)
    //TODO Apply tabColor
    virtuallyCenter.foreach(sheet.setVerticallyCenter)
    zoom.foreach(sheet.setZoom)
  }


}
