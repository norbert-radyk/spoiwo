package com.norbitltd.spoiwo.excel

import org.apache.poi.xssf.usermodel.XSSFSheet

object PrintSetup {

  val Default = apply()

}

case class PrintSetup private[Excel](copies : Option[Short] = None, draft : Option[Boolean] = None, fitHeight : Option[Short] = None,
                                     fitWidth : Option[Short] = None) {

  def withCopies(copies : Short) = copy(copies = Option(copies))
  def withDraft(draft : Boolean) = copy(draft = Option(draft))
  def withFitHeight(fitHeight : Short) = copy(fitHeight = Option(fitHeight))
  def withFitWidth(fitWidth : Short) = copy(fitWidth = Option(fitWidth))

  def applyTo(sheet : XSSFSheet) {
    if( this != PrintSetup.Default ){
      val ps = sheet.getPrintSetup
      copies.foreach(ps.setCopies(_))
      draft.foreach(ps.setDraft(_))
      fitHeight.foreach(ps.setFitHeight(_))
      fitWidth.foreach(ps.setFitWidth(_))
    }

    //TODO
  }
}
