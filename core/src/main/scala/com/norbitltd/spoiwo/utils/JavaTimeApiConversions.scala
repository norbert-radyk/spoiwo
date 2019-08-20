package com.norbitltd.spoiwo.utils

import java.util.Date
import java.time.{ZoneId, LocalDate => JLocalDate, LocalDateTime => JLocalDateTime}
import org.joda.time.{DateTime, LocalDate}

object JavaTimeApiConversions {

  implicit class RichJavaLocalDate(ld: JLocalDate) {
    def toDate: Date =
      new LocalDate(
        ld.getYear,
        ld.getMonthValue,
        ld.getDayOfMonth
      ).toDate
  }

  implicit class RichLocalDateTime(ldt: JLocalDateTime) {
    def toDate: Date =
      new DateTime(
        ldt.atZone(ZoneId.systemDefault()).toInstant.toEpochMilli
      ).toDate
  }
}
