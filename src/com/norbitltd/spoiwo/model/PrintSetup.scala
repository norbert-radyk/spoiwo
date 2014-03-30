package com.norbitltd.spoiwo.model

//TODO Enums for the following
import org.apache.poi.ss.usermodel.{PaperSize, PageOrder}

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

case class PrintSetup private(copies: Option[Short],
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

  def withCopies(copies: Short) = copy(copies = Option(copies))

  def withDraft(draft: Boolean) = copy(draft = Option(draft))

  def withFitHeight(fitHeight: Short) = copy(fitHeight = Option(fitHeight))

  def withFitWidth(fitWidth: Short) = copy(fitWidth = Option(fitWidth))

  def withFooterMargin(footerMargin: Double) = copy(footerMargin = Option(footerMargin))

  def withHeaderMargin(headerMargin: Double) = copy(headerMargin = Option(headerMargin))

  def withHResolution(hResolution: Short) = copy(hResolution = Option(hResolution))

  def withLandscape(landscape: Boolean) = copy(landscape = Option(landscape))

  def withLeftToRight(leftToRight: Boolean) = copy(leftToRight = Option(leftToRight))

  def withNoColor(noColor: Boolean) = copy(noColor = Option(noColor))

  def withNoOrientation(noOrientation: Boolean) = copy(noOrientation = Option(noOrientation))

  def withPageOrder(pageOrder: PageOrder) = copy(pageOrder = Option(pageOrder))

  def withPageStart(pageStart: Short) = copy(pageStart = Option(pageStart))

  def withPaperSize(paperSize: PaperSize) = copy(paperSize = Option(paperSize))

  def withScale(scale: Short) = copy(scale = Option(scale))

  def withUsePage(usePage: Boolean) = copy(usePage = Option(usePage))

  def withValidSettings(validSettings: Boolean) = copy(validSettings = Option(validSettings))

  def withVResolution(vResolution: Short) = copy(vResolution = Option(vResolution))

}
