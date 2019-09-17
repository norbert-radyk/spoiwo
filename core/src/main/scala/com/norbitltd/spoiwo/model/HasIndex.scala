package com.norbitltd.spoiwo.model

trait HasIndex[T] {
  def index(t: T): Option[Int]
}

object HasIndex {
  implicit object CellHasIndex extends HasIndex[Cell] {
    override def index(t: Cell): Option[Int] = t.index
  }
  implicit object RowHasIndex extends HasIndex[Row] {
    override def index(t: Row): Option[Int] = t.index
  }

  implicit class RichHasIndexSeq[T](indexed: Iterable[T])(implicit indexer: HasIndex[T]) {
    def maxIndex: Int = {
      indexed.foldLeft(0) { case (maxIdx, e) => math.max(maxIdx, indexer.index(e).getOrElse(maxIdx + 1))}
    }
  }
}
