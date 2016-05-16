package com.norbitltd.spoiwo.model

case class TableColumn(
  name: String,
  id: Long
) {

  def withName(name: String) =
    copy(name = name)

  def withId(id: Long) =
    copy(id = id)
}
