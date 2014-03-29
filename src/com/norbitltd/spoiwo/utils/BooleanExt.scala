package com.norbitltd.spoiwo.utils

trait BooleanExt {

  def toOption : Option[Boolean] = this match {
    case BooleanExt.Undefined => None
    case BooleanExt.False => Some(false)
    case BooleanExt.True => Some(true)
  }

}

object BooleanExt {

  implicit def boolean2NullableBoolean(value: Boolean): BooleanExt =
    if (value) True else False

  private[spoiwo] object Undefined extends BooleanExt

  object True extends BooleanExt

  object False extends BooleanExt

}


