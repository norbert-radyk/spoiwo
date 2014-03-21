package com.norbitltd.spoiwo.model

object SheetProperties extends Factory {

  private lazy val defaultAutoFilter = CellRange.None
  private lazy val defaultActiveCell = null
  private lazy val defaultAutoBreaks = true
  private lazy val defaultDefaultColumnWidth = -1
  private lazy val defaultDefaultRowHeight = Measure.Undefined

  private lazy val defaultDisplayFormulas = false
  private lazy val defaultDisplayGridLines = true
  private lazy val defaultDisplayGuts = true
  private lazy val defaultDisplayRowColHeadings = true
  private lazy val defaultDisplayZeros = true

  private lazy val defaultFitToPage = false
  private lazy val defaultForceFormulaRecalculation = false
  private lazy val defaultHorizontallyCenter = false
  private lazy val defaultPrintArea = CellRange.None
  private lazy val defaultPrintGridLines = false
  private lazy val defaultRightToLeft = false
  private lazy val defaultRowSumsBelow = true
  private lazy val defaultRowSumsRight = true
  private lazy val defaultSelected = false
  private lazy val defaultTabColor = -1
  private lazy val defaultVirtuallyCenter = false
  private lazy val defaultZoom = 100

  lazy val Default = SheetProperties()

  def apply(autoFilter: CellRange = defaultAutoFilter,
            activeCell: String = defaultActiveCell,
            autoBreaks: Boolean = defaultAutoBreaks,
            defaultColumnWidth: Int = defaultDefaultColumnWidth,
            defaultRowHeight: Measure = defaultDefaultRowHeight,
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

case class SheetProperties private[model](
                                           autoFilter: Option[CellRange],
                                           activeCell: Option[String],
                                           autoBreaks: Option[Boolean],
                                           defaultColumnWidth: Option[Int],
                                           defaultRowHeight: Option[Measure],
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

  def withoutAutoFilter(autoFilterRange: CellRange) =
    copy(autoFilter = None)

  def withActiveCell(activeCell: String) =
    copy(activeCell = Option(activeCell))

  def withoutActiveCell(activeCell: String) =
    copy(activeCell = None)

  def withAutoBreaks =
    copy(autoBreaks = Some(true))

  def withoutAutoBreaks =
    copy(autoBreaks = Some(false))

  def withDefaultColumnWidth(defaultColumnWidth: Int) =
    copy(defaultColumnWidth = Option(defaultColumnWidth))

  def withoutDefaultColumnWidth(defaultColumnWidth: Int) =
    copy(defaultColumnWidth = None)

  def withDefaultRowHeight(defaultRowHeight: Measure) =
    copy(defaultRowHeight = Option(defaultRowHeight))

  def withoutDefaultRowHeight(defaultRowHeight: Short) =
    copy(defaultRowHeight = None)

  def withDisplayFormulas =
    copy(displayFormulas = Some(true))

  def withoutDisplayFormulas =
    copy(displayFormulas = Some(false))

  def withDisplayGridLines =
    copy(displayGridLines = Some(true))

  def withoutDisplayGridLines =
    copy(displayGridLines = Some(false))

  def withDisplayGuts =
    copy(displayGuts = Some(true))

  def withoutDisplayGuts =
    copy(displayGuts = Some(false))

  def withDisplayRowColHeadings =
    copy(displayRowColHeadings = Some(true))

  def withoutDisplayRowColHeadings =
    copy(displayRowColHeadings = Some(false))

  def withDisplayZeros =
    copy(displayZeros = Some(true))

  def withoutDisplayZeros =
    copy(displayZeros = Some(false))

  def withFitToPage =
    copy(fitToPage = Some(true))

  def withoutFitToPage =
    copy(fitToPage = Some(false))

  def withForceFormulaRecalculation =
    copy(forceFormulaRecalculation = Some(true))

  def withoutForceFormulaRecalculation =
    copy(forceFormulaRecalculation = Some(false))

  def withHorizontallyCenter =
    copy(horizontallyCenter = Some(true))

  def withoutHorizontallyCenter =
    copy(horizontallyCenter = Some(false))

  def withPrintArea(printArea: CellRange) =
    copy(printArea = Option(printArea))

  def withoutPrintArea(printArea: CellRange) =
    copy(printArea = None)

  def withPrintGridLines =
    copy(printGridLines = Some(true))

  def withoutPrintGridLines =
    copy(printGridLines = Some(false))

  def withRightToLeft =
    copy(rightToLeft = Some(true))

  def withoutRightToLeft =
    copy(rightToLeft = Some(false))

  def withRowSumsBelow =
    copy(rowSumsBelow = Some(true))

  def withoutRowSumsBelow =
    copy(rowSumsBelow = Some(false))

  def withRowSumsRight =
    copy(rowSumsRight = Some(true))

  def withoutRowSumsRight =
    copy(rowSumsRight = Some(false))

  def withSelected =
    copy(selected = Some(true))

  def withoutSelected =
    copy(selected = Some(false))

  def withTabColor(tabColor: Int) =
    copy(tabColor = Option(tabColor))

  def withoutTabColor =
    copy(tabColor = None)

  def withVirtuallyCenter =
    copy(virtuallyCenter = Option(true))

  def withoutVirtuallyCenter =
    copy(virtuallyCenter = Option(false))

  def withZoom(zoom: Int) =
    copy(zoom = Option(zoom))

  def withoutZoom =
    copy(zoom = None)

}
