package com.norbitltd.spoiwo.utils

import java.time.{ LocalDate => JLocalDate, LocalDateTime => JLocalDateTime}
import org.scalatest.FlatSpec
import org.joda.time.{DateTime, LocalDate}
import JavaTimeApiConversions._

class JavaTimeApiConversionsSpec extends FlatSpec {

  "Conversion from Java LocalDate to Date" should "produce same Date as from corresponding Joda LocalDate" in {

    test(JLocalDate.of(2011, 6, 13), new LocalDate(2011, 6, 13))
    test(JLocalDate.of(2011, 11, 13), new LocalDate(2011, 11, 13))

    def test(jld: JLocalDate, ld: LocalDate): Unit = {
      assert(jld.toDate === ld.toDate)
    }
  }

  "Conversion from Java LocalDateTime to Date" should "produce same Date as from corresponding Joda DateTime" in {

    test(JLocalDateTime.of(2011, 6, 13, 15, 30, 10, 5000000), new DateTime(2011, 6, 13, 15, 30, 10, 5))
    test(JLocalDateTime.of(2011, 11, 13, 15, 30, 10, 999999999), new DateTime(2011, 11, 13, 15, 30, 10, 999))

    def test(jldt: JLocalDateTime, ld: DateTime): Unit = {
      assert(jldt.toDate === ld.toDate)
    }
  }
}
