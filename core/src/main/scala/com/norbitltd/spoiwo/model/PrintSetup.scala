package com.norbitltd.spoiwo.model

import com.norbitltd.spoiwo.model.enums.{PaperSize, PageOrder}

object PrintSetup {

  val Default = PrintSetup()

  def apply(copies: java.lang.Integer = null,
            draft: java.lang.Boolean = null,
            fitHeight: java.lang.Integer = null,
            fitWidth: java.lang.Integer = null,
            footerMargin: java.lang.Double = null,
            headerMargin: java.lang.Double = null,
            hResolution: java.lang.Integer = null,
            landscape: java.lang.Boolean = null,
            leftToRight: java.lang.Boolean = null,
            noColor: java.lang.Boolean = null,
            noOrientation: java.lang.Boolean = null,
            pageOrder: PageOrder = null,
            pageStart: java.lang.Integer = null,
            paperSize: PaperSize = null,
            scale: java.lang.Integer = null,
            usePage: java.lang.Boolean = null,
            validSettings: java.lang.Boolean = null,
            vResolution: java.lang.Integer = null): PrintSetup =
    new PrintSetup(
      copies = Option(copies).map(_.shortValue),
      draft = Option(draft).map(_.booleanValue),
      fitHeight = Option(fitHeight).map(_.shortValue),
      fitWidth = Option(fitWidth).map(_.shortValue),
      footerMargin = Option(footerMargin).map(_.shortValue),
      headerMargin = Option(headerMargin).map(_.shortValue),
      hResolution = Option(hResolution).map(_.shortValue),
      landscape = Option(landscape).map(_.booleanValue),
      leftToRight = Option(leftToRight).map(_.booleanValue),
      noColor = Option(noColor).map(_.booleanValue),
      noOrientation = Option(noOrientation).map(_.booleanValue),
      pageOrder = Option(pageOrder),
      pageStart = Option(pageStart).map(_.shortValue),
      paperSize = Option(paperSize),
      scale = Option(scale).map(_.shortValue),
      usePage = Option(usePage).map(_.booleanValue),
      validSettings = Option(validSettings).map(_.booleanValue),
      vResolution = Option(vResolution).map(_.shortValue)
    )

}

case class PrintSetup private (copies: Option[Short],
                               draft: Option[Boolean],
                               fitHeight: Option[Short],
                               fitWidth: Option[Short],
                               footerMargin: Option[Double],
                               headerMargin: Option[Double],
                               hResolution: Option[Short],
                               landscape: Option[Boolean],
                               leftToRight: Option[Boolean],
                               noColor: Option[Boolean],
                               noOrientation: Option[Boolean],
                               pageOrder: Option[PageOrder],
                               pageStart: Option[Short],
                               paperSize: Option[PaperSize],
                               scale: Option[Short],
                               usePage: Option[Boolean],
                               validSettings: Option[Boolean],
                               vResolution: Option[Short]) {

  def withCopies(copies: Short): PrintSetup = copy(copies = Option(copies))

  def withDraft(draft: Boolean): PrintSetup = copy(draft = Option(draft))

  def withFitHeight(fitHeight: Short): PrintSetup = copy(fitHeight = Option(fitHeight))

  def withFitWidth(fitWidth: Short): PrintSetup = copy(fitWidth = Option(fitWidth))

  def withFooterMargin(footerMargin: Double): PrintSetup = copy(footerMargin = Option(footerMargin))

  def withHeaderMargin(headerMargin: Double): PrintSetup = copy(headerMargin = Option(headerMargin))

  def withHResolution(hResolution: Short): PrintSetup = copy(hResolution = Option(hResolution))

  def withLandscape(landscape: Boolean): PrintSetup = copy(landscape = Option(landscape))

  def withLeftToRight(leftToRight: Boolean): PrintSetup = copy(leftToRight = Option(leftToRight))

  def withNoColor(noColor: Boolean): PrintSetup = copy(noColor = Option(noColor))

  def withNoOrientation(noOrientation: Boolean): PrintSetup = copy(noOrientation = Option(noOrientation))

  def withPageOrder(pageOrder: PageOrder): PrintSetup = copy(pageOrder = Option(pageOrder))

  def withPageStart(pageStart: Short): PrintSetup = copy(pageStart = Option(pageStart))

  def withPaperSize(paperSize: PaperSize): PrintSetup = copy(paperSize = Option(paperSize))

  def withScale(scale: Short): PrintSetup = copy(scale = Option(scale))

  def withUsePage(usePage: Boolean): PrintSetup = copy(usePage = Option(usePage))

  def withValidSettings(validSettings: Boolean): PrintSetup = copy(validSettings = Option(validSettings))

  def withVResolution(vResolution: Short): PrintSetup = copy(vResolution = Option(vResolution))

}
