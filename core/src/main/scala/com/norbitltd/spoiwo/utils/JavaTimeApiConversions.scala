package com.norbitltd.spoiwo.utils

import java.util.Date
import java.time.{ZoneId, LocalDate, LocalDateTime}

object JavaTimeApiConversions {

  implicit class RichJavaLocalDate(ld: LocalDate) {
    def toDate: Date = Date.from(ld.atStartOfDay().atZone(ZoneId.systemDefault()).toInstant)
  }

  implicit class RichLocalDateTime(ldt: LocalDateTime) {
    def toDate: Date = Date.from(ldt.atZone(ZoneId.systemDefault()).toInstant)
  }
}
