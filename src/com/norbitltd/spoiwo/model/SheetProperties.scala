package com.norbitltd.spoiwo.model

object SheetProperties extends Factory {

  private lazy val defaultAutoFilter = CellRange.None
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
  private lazy val defaultPrintArea = CellRange.None
  private lazy val defaultPrintGridLines = false
  private lazy val defaultRightToLeft = false
  private lazy val defaultRowSumsBelow = defaultPOISheet.getRowSumsBelow
  private lazy val defaultRowSumsRight = defaultPOISheet.getRowSumsRight
  private lazy val defaultSelected = false
  private lazy val defaultTabColor = -1
  private lazy val defaultVirtuallyCenter = defaultPOISheet.getVerticallyCenter
  private lazy val defaultZoom = 100

  lazy val Default = SheetProperties()

  def apply(autoFilter: CellRange = defaultAutoFilter,
            activeCell: String = defaultActiveCell,
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
            printArea: CellRange = defaultPrintArea,
            printGridLines: Boolean = defaultPrintGridLines,
            rightToLeft: Boolean = defaultRightToLeft,
            rowSumsBelow: Boolean = defaultRowSumsBelow,
            rowSumsRight: Boolean = defaultRowSumsRight,
            selected: Boolean = defaultSelected,
            tabColor: Int = defaultTabColor,
            virtuallyCenter: Boolean = defaultVirtuallyCenter,
            zoom: Int = defaultZoom): SheetProperties = SheetProperties(
    wrap(autoFilter, autoFilter),
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
    wrap(printArea, defaultPrintArea),
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

/**
 * Represents the complete set of Apache POI sheet properties.
 * The following methods are not supported:
 * setRepeatingColumns and setRepeatingRows
 */
case class SheetProperties private[model](
                                           autoFilter: Option[CellRange],
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
                                           printArea: Option[CellRange],
                                           printGridLines: Option[Boolean],
                                           rightToLeft: Option[Boolean],
                                           rowSumsBelow: Option[Boolean],
                                           rowSumsRight: Option[Boolean],
                                           selected: Option[Boolean],
                                           tabColor: Option[Int],
                                           virtuallyCenter: Option[Boolean],
                                           zoom: Option[Int]) {

  def withAutoFilter(autoFilterRange: CellRange) =
    copy(autoFilter = Option(autoFilterRange))

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

  def withPrintArea(printArea: CellRange) =
    copy(printArea = Option(printArea))

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

  def withTabColor(tabColor: Int) =
    copy(tabColor = Option(tabColor))

  def withVirtuallyCenter(virtuallyCenter: Boolean) =
    copy(virtuallyCenter = Option(virtuallyCenter))

  def withZoom(zoom: Int) =
    copy(zoom = Option(zoom))

}
