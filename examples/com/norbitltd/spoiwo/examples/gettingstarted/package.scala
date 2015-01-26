package com.norbitltd.spoiwo.examples.gettingstarted

import java.io.File
import scala.swing.FileChooser

package object utils {
  
  def returnOrAskForSaveLocation(args:Array[String]):String = 
    if(args.length >= 1) {
      args(0)
    } else {
      askWhereToSave().getOrElse(
          throw new IllegalStateException("Please specify where to save the sample spreadsheet")
      )
    }
  
  private def askWhereToSave():Option[String] = {
    val chooser = new FileChooser(new File("."))
    chooser.title = "Save sample to"
    val result = chooser.showSaveDialog(null)
    if (result == FileChooser.Result.Approve) {
      Some(chooser.selectedFile.getAbsolutePath)
    } else None
  }
}