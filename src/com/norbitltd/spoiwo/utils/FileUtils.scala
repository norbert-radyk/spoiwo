package com.norbitltd.spoiwo.utils

import java.io.File

object FileUtils {

  private def printToFile(f: java.io.File)(op: java.io.PrintWriter => Unit) {
    val p = new java.io.PrintWriter(f)
    try { op(p) } finally { p.flush(); p.close() }
  }

  def write(fileName : String, content : String) {
    printToFile(new File(fileName)) {
      p => p.write(content)
    }
  }

}
