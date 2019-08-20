package com.norbitltd.spoiwo.model

case class TableColumn(
    name: String,
    id: Long
) {

  def withName(name: String): TableColumn =
    copy(name = name)

  def withId(id: Long): TableColumn =
    copy(id = id)
}
