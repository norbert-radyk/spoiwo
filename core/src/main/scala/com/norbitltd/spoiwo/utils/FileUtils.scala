package com.norbitltd.spoiwo.utils

import java.io.{File, PrintWriter}

object FileUtils {

  def write(fileName: String, content: String): Unit = {
    val file = new File(fileName)
    val writer = new PrintWriter(file)

    try {
      writer.write(content)
      writer.flush()
    } finally {
      writer.close()
    }
  }

}
