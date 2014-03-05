package com.norbitltd.spoiwo.model

object Column extends Factory {

  private lazy val defaultIndex = -1
  private lazy val defaultAutoSized = false
  private lazy val defaultBreak = false
  private lazy val defaultGroupCollapsed = false
  private lazy val defaultHidden = false
  private lazy val defaultStyle = CellStyle.Default
  private lazy val defaultWidth = defaultPOISheet.getDefaultColumnWidth

  val Default = Column()

  def apply(
             index: Int = defaultIndex,
             autoSized: Boolean = defaultAutoSized,
             break: Boolean = defaultBreak,
             groupCollapsed: Boolean = defaultGroupCollapsed,
             hidden: Boolean = defaultHidden,
             style: CellStyle = defaultStyle,
             width: Int = defaultWidth): Column =
    Column(
      index = wrap(index, defaultIndex),
      autoSized = wrap(autoSized, defaultAutoSized),
      break = wrap(break, defaultBreak),
      groupCollapsed = wrap(groupCollapsed, defaultGroupCollapsed),
      hidden = wrap(hidden, defaultHidden),
      style = wrap(style, defaultStyle),
      width = wrap(width, defaultWidth)
    )
}

case class Column private[model](index: Option[Int],
                              autoSized: Option[Boolean],
                              break: Option[Boolean],
                              groupCollapsed: Option[Boolean],
                              hidden: Option[Boolean],
                              style: Option[CellStyle],
                              width: Option[Int]) {

  def withIndex(index: Int) =
    copy(index = Option(index))

  def withAutoSized(autoSized: Boolean) =
    copy(autoSized = Option(autoSized))

  def withBreak() =
    copy(break = Option(true))

  def withGroupCollapsed(groupCollapsed: Boolean) =
    copy(groupCollapsed = Option(groupCollapsed))

  def withHidden(hidden: Boolean) =
    copy(hidden = Option(hidden))

  def withStyle(style: CellStyle) =
    copy(style = Option(style))

  def withWidth(width: Int) =
    copy(width = Option(width))
}
