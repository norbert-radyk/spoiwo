package com.norbitltd.spoiwo.model

object Column {

  def apply(index: java.lang.Integer = null,
            autoSized: java.lang.Boolean = null,
            break: java.lang.Boolean = null,
            groupCollapsed: java.lang.Boolean = null,
            hidden: java.lang.Boolean = null,
            style: CellStyle = null,
            width: Width = null): Column =
    Column(
      index = Option(index).map(_.intValue),
      autoSized = Option(autoSized).map(_.booleanValue),
      break = Option(break).map(_.booleanValue),
      groupCollapsed = Option(groupCollapsed).map(_.booleanValue),
      hidden = Option(hidden).map(_.booleanValue),
      style = Option(style),
      width = Option(width)
    )
}

case class Column private[model] (index: Option[Int],
                                  autoSized: Option[Boolean],
                                  break: Option[Boolean],
                                  groupCollapsed: Option[Boolean],
                                  hidden: Option[Boolean],
                                  style: Option[CellStyle],
                                  width: Option[Width]) {

  override def toString: String =
    "Column(" + List(
      index.map("index=" + _),
      autoSized.map("auto sized=" + _),
      break.map("break=" + _),
      groupCollapsed.map("group collapsed=" + _),
      hidden.map("hidden=" + _),
      width.map("width=" + _)
    ).flatten.mkString(", ") + ")"

  def withIndex(index: Int): Column =
    copy(index = Option(index))

  def withoutIndex: Column =
    copy(index = None)

  def withAutoSized: Column =
    copy(autoSized = Some(true))

  def withoutAutoSized: Column =
    copy(autoSized = Some(false))

  def withBreak: Column =
    copy(break = Some(true))

  def withoutBreak: Column =
    copy(break = Some(false))

  def withGroupCollapsed: Column =
    copy(groupCollapsed = Some(true))

  def withoutGroupCollapsed: Column =
    copy(groupCollapsed = Some(false))

  def withHidden: Column =
    copy(hidden = Some(true))

  def withoutHidden: Column =
    copy(hidden = Some(false))

  def withStyle(style: CellStyle): Column =
    copy(style = Option(style))

  def withoutStyle: Column =
    copy(style = None)

  def withWidth(width: Width): Column =
    copy(width = Option(width))

  def withoutWidth: Column =
    copy(width = None)
}
