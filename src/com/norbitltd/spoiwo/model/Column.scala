package com.norbitltd.spoiwo.model

object Column extends Factory {

  private lazy val defaultIndex = -1
  private lazy val defaultAutoSized = false
  private lazy val defaultBreak = false
  private lazy val defaultGroupCollapsed = false
  private lazy val defaultHidden = false
  private lazy val defaultStyle = CellStyle.Default
  private lazy val defaultWidth = Width.Undefined

  val Default = Column()

  def apply(
             index: Int = defaultIndex,
             autoSized: Boolean = defaultAutoSized,
             break: Boolean = defaultBreak,
             groupCollapsed: Boolean = defaultGroupCollapsed,
             hidden: Boolean = defaultHidden,
             style: CellStyle = defaultStyle,
             width: Width = defaultWidth): Column =
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
                              width: Option[Width]) {

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
