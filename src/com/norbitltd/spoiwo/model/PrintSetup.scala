package com.norbitltd.spoiwo.model

//TODO Enums for the following
import org.apache.poi.ss.usermodel.{PaperSize, PageOrder}

object PrintSetup extends Factory {

  private lazy val defaultCopies = defaultPOIPrintSetup.getCopies
  private lazy val defaultDraft = defaultPOIPrintSetup.getDraft
  private lazy val defaultFitHeight = defaultPOIPrintSetup.getFitHeight
  private lazy val defaultFitWidth = defaultPOIPrintSetup.getFitWidth
  private lazy val defaultFooterMargin = defaultPOIPrintSetup.getFooterMargin
  private lazy val defaultHeaderMargin = defaultPOIPrintSetup.getHeaderMargin
  private lazy val defaultHResolution = defaultPOIPrintSetup.getHResolution
  private lazy val defaultLandscape = defaultPOIPrintSetup.getLandscape
  private lazy val defaultLeftToRight = defaultPOIPrintSetup.getLeftToRight
  private lazy val defaultNoColor = defaultPOIPrintSetup.getNoColor
  private lazy val defaultNoOrientation = defaultPOIPrintSetup.getNoOrientation
  private lazy val defaultPageOrder = defaultPOIPrintSetup.getPageOrder
  private lazy val defaultPageStart = defaultPOIPrintSetup.getPageStart
  private lazy val defaultPaperSize = defaultPOIPrintSetup.getPaperSizeEnum
  private lazy val defaultScale = defaultPOIPrintSetup.getScale
  private lazy val defaultUsePage = defaultPOIPrintSetup.getUsePage
  private lazy val defaultValidSettings = defaultPOIPrintSetup.getValidSettings
  private lazy val defaultVResolution = defaultPOIPrintSetup.getVResolution

  val Default = apply()

  def apply(copies: Short = defaultCopies,
            draft: Boolean = defaultDraft,
            fitHeight: Short = defaultFitHeight,
            fitWidth: Short = defaultFitWidth,
            footerMargin: Double = defaultFooterMargin,
            headerMargin: Double = defaultHeaderMargin,
            hResolution: Short = defaultHResolution,
            landscape: Boolean = defaultLandscape,
            leftToRight: Boolean = defaultLeftToRight,
            noColor: Boolean = defaultNoColor,
            noOrientation: Boolean = defaultNoOrientation,
            pageOrder: PageOrder = defaultPageOrder,
            pageStart: Short = defaultPageStart,
            paperSize: PaperSize = defaultPaperSize,
            scale: Short = defaultScale,
            usePage: Boolean = defaultUsePage,
            validSettings: Boolean = defaultValidSettings,
            vResolution: Short = defaultVResolution): PrintSetup =
    PrintSetup(
      copies = wrap(copies, defaultCopies),
      draft = wrap(draft, defaultDraft),
      fitHeight = wrap(fitHeight, defaultFitHeight),
      fitWidth = wrap(fitWidth, defaultFitWidth),
      footerMargin = wrap(footerMargin, defaultFooterMargin),
      headerMargin = wrap(headerMargin, defaultHeaderMargin),
      hResolution = wrap(hResolution, defaultHResolution),
      landscape = wrap(landscape, defaultLandscape),
      leftToRight = wrap(leftToRight, defaultLeftToRight),
      noColor = wrap(noColor, defaultNoColor),
      noOrientation = wrap(noOrientation, defaultNoOrientation),
      pageOrder = wrap(pageOrder, defaultPageOrder),
      pageStart = wrap(pageStart, defaultPageStart),
      paperSize = wrap(paperSize, defaultPaperSize),
      scale = wrap(scale, defaultScale),
      usePage = wrap(usePage, defaultUsePage),
      validSettings = wrap(validSettings, defaultValidSettings),
      vResolution = wrap(vResolution, defaultVResolution)
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
