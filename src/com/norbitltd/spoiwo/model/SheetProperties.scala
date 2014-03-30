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
            tabColor: java.lang.Integer = null,
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
    Option(tabColor).map(_.intValue),
    Option(virtuallyCenter).map(_.booleanValue),
    Option(zoom).map(_.intValue)
  )
}

case class SheetProperties private[model](
                                           autoFilter: Option[CellRange],
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

  def withDefaultRowHeight(defaultRowHeight: Height) =
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
