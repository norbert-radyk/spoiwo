package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.ss.usermodel.{PaperSize, PageOrder}

object PrintSetup {

  val Default = apply()

}

case class PrintSetup private[Excel](copies: Option[Short] = None,
                                     draft: Option[Boolean] = None,
                                     fitHeight: Option[Short] = None,
                                     fitWidth: Option[Short] = None,
                                     footerMargin: Option[Double] = None,
                                     headerMargin: Option[Double] = None,
                                     resolution: Option[Short] = None,
                                     landscape: Option[Boolean] = None,
                                     leftToRight: Option[Boolean] = None,
                                     noColor: Option[Boolean] = None,
                                     noOrientation: Option[Boolean] = None,
                                     pageOrder: Option[PageOrder] = None,
                                     pageStart: Option[Short] = None,
                                     paperSize: Option[PaperSize] = None,
                                     scale: Option[Short] = None,
                                     usePage: Option[Boolean] = None,
                                     validSettings: Option[Boolean] = None,
                                     vResolution: Option[Short] = None) {

  def withCopies(copies: Short) = copy(copies = Option(copies))
  def withDraft(draft: Boolean) = copy(draft = Option(draft))
  def withFitHeight(fitHeight: Short) = copy(fitHeight = Option(fitHeight))
  def withFitWidth(fitWidth: Short) = copy(fitWidth = Option(fitWidth))
  def withFooterMargin(footerMargin: Double) = copy(footerMargin = Option(footerMargin))
  def withHeaderMargin(headerMargin: Double) = copy(headerMargin = Option(headerMargin))
  def withResolution(resolution: Short) = copy(resolution = Option(resolution))
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


  def applyTo(sheet: XSSFSheet) {
    if (this != PrintSetup.Default) {
      val ps = sheet.getPrintSetup
      copies.foreach(ps.setCopies(_))
      draft.foreach(ps.setDraft(_))
      fitHeight.foreach(ps.setFitHeight(_))
      fitWidth.foreach(ps.setFitWidth(_))
      footerMargin.foreach(ps.setFooterMargin(_))
      headerMargin.foreach(ps.setHeaderMargin(_))
      resolution.foreach(ps.setHResolution(_))
      landscape.foreach(ps.setLandscape(_))
      leftToRight.foreach(ps.setLeftToRight(_))
      noColor.foreach(ps.setNoColor(_))
      noOrientation.foreach(ps.setNoOrientation(_))
      pageOrder.foreach(ps.setPageOrder(_))
      pageStart.foreach(ps.setPageStart(_))
      paperSize.foreach(ps.setPaperSize(_))
      scale.foreach(ps.setScale(_))
      usePage.foreach(ps.setUsePage(_))
      validSettings.foreach(ps.setValidSettings(_))
      vResolution.foreach(ps.setVResolution(_))
    }
  }
}
