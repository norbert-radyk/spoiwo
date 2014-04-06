package com.norbitltd.spoiwo.model

object Column {

  def apply(
             index: java.lang.Integer = null,
             autoSized: java.lang.Boolean = null,
             break: java.lang.Boolean = null,
             groupCollapsed: java.lang.Boolean = null,
             hidden: java.lang.Boolean = null,
             style : CellStyle = null,
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

case class Column private[model](index: Option[Int],
                                 autoSized: Option[Boolean],
                                 break: Option[Boolean],
                                 groupCollapsed: Option[Boolean],
                                 hidden: Option[Boolean],
                                 style: Option[CellStyle],
                                 width: Option[Width]) {

  override def toString = "Column(" + List(
    index.map("index=" + _),
    autoSized.map("auto sized=" + _),
    break.map("break=" + _),
    groupCollapsed.map("group collapsed=" + _),
    hidden.map("hidden=" + _),
    //style.map("style = " + _),
    width.map("width=" + _)
  ).flatten.mkString(", ") + ")"

  def withIndex(index: Int) =
    copy(index = Option(index))

  def withoutIndex =
    copy(index = None)

  def withAutoSized(autoSized: Boolean) =
    copy(autoSized = Some(true))

  def withoutAutoSized =
    copy(autoSized = Some(false))

  def withBreak =
    copy(break = Some(true))

  def withoutBreak =
    copy(break = Some(false))

  def withGroupCollapsed =
    copy(groupCollapsed = Some(true))

  def withoutGroupCollapsed =
    copy(groupCollapsed = Some(false))

  def withHidden =
    copy(hidden = Some(true))

  def withoutHidden =
    copy(hidden = Some(false))

  def withStyle(style: CellStyle) =
    copy(style = Option(style))

  def withoutStyle =
    copy(style = None)

  def withWidth(width: Width) =
    copy(width = Option(width))

  def withoutWidth =
    copy(width = None)
}
