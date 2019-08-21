package com.norbitltd.spoiwo.utils

import java.time.{LocalDate => JLocalDate, LocalDateTime => JLocalDateTime}

import org.scalatest.{FlatSpec, Matchers}
import org.joda.time.{DateTime, LocalDate}
import JavaTimeApiConversions._

class JavaTimeApiConversionsSpec extends FlatSpec with Matchers {

  "Conversion from Java LocalDate to Date" should "produce same Date as from corresponding Joda LocalDate" in {
    JLocalDate.of(2011, 6, 13).toDate shouldBe new LocalDate(2011, 6, 13).toDate
    JLocalDate.of(2011, 11, 13).toDate shouldBe new LocalDate(2011, 11, 13).toDate
  }

  "Conversion from Java LocalDateTime to Date" should "produce same Date as from corresponding Joda DateTime" in {
    JLocalDateTime.of(2011, 6, 13, 15, 30, 10, 5000000).toDate shouldBe new DateTime(2011, 6, 13, 15, 30, 10, 5).toDate
    JLocalDateTime.of(2011, 11, 13, 15, 30, 10, 999999999).toDate shouldBe new DateTime(2011, 11, 13, 15, 30, 10,
      999).toDate
  }
}
