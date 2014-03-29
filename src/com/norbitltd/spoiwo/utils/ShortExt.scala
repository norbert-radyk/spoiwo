package com.norbitltd.spoiwo.utils

sealed trait ShortExt {

  def toOption: Option[Short] = this match {
    case ShortExt.Undefined => None
    case ShortExt.Defined(x) => Option(x)
  }

}

object ShortExt {

  implicit def int2ShortExtension(value : Int) : ShortExt =
    short2ShortExtension(value.toShort)

  implicit def short2ShortExtension(value: Short): ShortExt =
    Defined(value)

  private[spoiwo] object Undefined extends ShortExt

  case class Defined(value: Short) extends ShortExt

}
