package com.norbitltd.spoiwo.model

object SheetProperties {

  lazy val Default = SheetProperties()

  def apply(autoFilter: CellRange = null,
            activeCell: String = null,
            autoBreaks: java.lang.Boolean = null,
            defaultColumnWidth: java.lang.Integer = null,
            defaultRowHeight: Height = null,
            displayFormulas: java.lang.Boolean = null,
            displayGridLines: java.lang.Boolean = null,
            displayGuts: java.lang.Boolean = null,
            displayRowColHeadings: java.lang.Boolean = null,
            displayZeros: java.lang.Boolean = null,
            fitToPage: java.lang.Boolean = null,
            forceFormulaRecalculation: java.lang.Boolean = null,
            horizontallyCenter: java.lang.Boolean = null,
            printArea: CellRange = null,
            printGridLines: java.lang.Boolean = null,
            rightToLeft: java.lang.Boolean = null,
            rowSumsBelow: java.lang.Boolean = null,
            rowSumsRight: java.lang.Boolean = null,
            selected: java.lang.Boolean = null,
            tabColor: Color = null,
            virtuallyCenter: java.lang.Boolean = null,
            zoom: java.lang.Integer = null): SheetProperties = SheetProperties(
    Option(autoFilter),
    Option(activeCell),
    Option(autoBreaks).map(_.booleanValue),
    Option(defaultColumnWidth).map(_.intValue),
    Option(defaultRowHeight),
    Option(displayFormulas).map(_.booleanValue),
    Option(displayGridLines).map(_.booleanValue),
    Option(displayGuts).map(_.booleanValue),
    Option(displayRowColHeadings).map(_.booleanValue),
    Option(displayZeros).map(_.booleanValue),
    Option(fitToPage).map(_.booleanValue),
    Option(forceFormulaRecalculation).map(_.booleanValue),
    Option(horizontallyCenter).map(_.booleanValue),
    Option(printArea),
    Option(printGridLines).map(_.booleanValue),
    Option(rightToLeft).map(_.booleanValue),
    Option(rowSumsBelow).map(_.booleanValue),
    Option(rowSumsRight).map(_.booleanValue),
    Option(selected).map(_.booleanValue),
    Option(tabColor),
    Option(virtuallyCenter).map(_.booleanValue),
    Option(zoom).map(_.intValue)
  )
}

case class SheetProperties private[model] (autoFilter: Option[CellRange],
                                           activeCell: Option[String],
                                           autoBreaks: Option[Boolean],
                                           defaultColumnWidth: Option[Int],
                                           defaultRowHeight: Option[Height],
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
                                           tabColor: Option[Color],
                                           virtuallyCenter: Option[Boolean],
                                           zoom: Option[Int]) {

  def withAutoFilter(autoFilterRange: CellRange): SheetProperties =
    copy(autoFilter = Option(autoFilterRange))

  def withoutAutoFilter: SheetProperties =
    copy(autoFilter = None)

  def withActiveCell(activeCell: String): SheetProperties =
    copy(activeCell = Option(activeCell))

  def withoutActiveCell: SheetProperties =
    copy(activeCell = None)

  def withAutoBreaks: SheetProperties =
    copy(autoBreaks = Some(true))

  def withoutAutoBreaks: SheetProperties =
    copy(autoBreaks = Some(false))

  def withDefaultColumnWidth(defaultColumnWidth: Int): SheetProperties =
    copy(defaultColumnWidth = Option(defaultColumnWidth))

  def withoutDefaultColumnWidth: SheetProperties =
    copy(defaultColumnWidth = None)

  def withDefaultRowHeight(defaultRowHeight: Height): SheetProperties =
    copy(defaultRowHeight = Option(defaultRowHeight))

  def withoutDefaultRowHeight: SheetProperties =
    copy(defaultRowHeight = None)

  def withDisplayFormulas: SheetProperties =
    copy(displayFormulas = Some(true))

  def withoutDisplayFormulas: SheetProperties =
    copy(displayFormulas = Some(false))

  def withDisplayGridLines: SheetProperties =
    copy(displayGridLines = Some(true))

  def withoutDisplayGridLines: SheetProperties =
    copy(displayGridLines = Some(false))

  def withDisplayGuts: SheetProperties =
    copy(displayGuts = Some(true))

  def withoutDisplayGuts: SheetProperties =
    copy(displayGuts = Some(false))

  def withDisplayRowColHeadings: SheetProperties =
    copy(displayRowColHeadings = Some(true))

  def withoutDisplayRowColHeadings: SheetProperties =
    copy(displayRowColHeadings = Some(false))

  def withDisplayZeros: SheetProperties =
    copy(displayZeros = Some(true))

  def withoutDisplayZeros: SheetProperties =
    copy(displayZeros = Some(false))

  def withFitToPage: SheetProperties =
    copy(fitToPage = Some(true))

  def withoutFitToPage: SheetProperties =
    copy(fitToPage = Some(false))

  def withForceFormulaRecalculation: SheetProperties =
    copy(forceFormulaRecalculation = Some(true))

  def withoutForceFormulaRecalculation: SheetProperties =
    copy(forceFormulaRecalculation = Some(false))

  def withHorizontallyCenter: SheetProperties =
    copy(horizontallyCenter = Some(true))

  def withoutHorizontallyCenter: SheetProperties =
    copy(horizontallyCenter = Some(false))

  def withPrintArea(printArea: CellRange): SheetProperties =
    copy(printArea = Option(printArea))

  def withoutPrintArea: SheetProperties =
    copy(printArea = None)

  def withPrintGridLines: SheetProperties =
    copy(printGridLines = Some(true))

  def withoutPrintGridLines: SheetProperties =
    copy(printGridLines = Some(false))

  def withRightToLeft: SheetProperties =
    copy(rightToLeft = Some(true))

  def withoutRightToLeft: SheetProperties =
    copy(rightToLeft = Some(false))

  def withRowSumsBelow: SheetProperties =
    copy(rowSumsBelow = Some(true))

  def withoutRowSumsBelow: SheetProperties =
    copy(rowSumsBelow = Some(false))

  def withRowSumsRight: SheetProperties =
    copy(rowSumsRight = Some(true))

  def withoutRowSumsRight: SheetProperties =
    copy(rowSumsRight = Some(false))

  def withSelected: SheetProperties =
    copy(selected = Some(true))

  def withoutSelected: SheetProperties =
    copy(selected = Some(false))

  def withTabColor(tabColor: Color): SheetProperties =
    copy(tabColor = Option(tabColor))

  def withoutTabColor: SheetProperties =
    copy(tabColor = None)

  def withVirtuallyCenter: SheetProperties =
    copy(virtuallyCenter = Option(true))

  def withoutVirtuallyCenter: SheetProperties =
    copy(virtuallyCenter = Option(false))

  def withZoom(zoom: Int): SheetProperties =
    copy(zoom = Option(zoom))

  def withoutZoom: SheetProperties =
    copy(zoom = None)

}
