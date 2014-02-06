package com.norbitltd.spoiwo.excel.util

object CodeGen {

  def main(args: Array[String]) {
    val params = List("top" -> "Double", "bottom" -> "Double", "right" -> "Double", "left" -> "Double", "header" -> "Double", "footer" -> "Double")
    println(createApplyMethod("Margins", params))
    println(createClass("Margins", params))
  }

  def createApplyMethod(objectName: String, values: List[(String, String)]): String = {
    "def apply(" +
      values.map {
        case (valueName, valueType) => "%s: %s = default%s,"
          .format(valueName, valueType, valueName.capitalize)
      }.mkString("\n") +
      "): %s = %s(".format(objectName, objectName) +
     values.map { case(valueName, valueType) => "wrap(%s, default%s),".format(valueName, valueName.capitalize) }.mkString("\n") +
    "\n)"
  }

  def createClass(objectName: String, values: List[(String, String)]) : String = {
    "case class %s private(\n".format(objectName) +
    values.map { case(valueName, valueType) => "%s: Option[%s],".format(valueName, valueType)}.mkString("\n")  +
    ") {\n" +
    values.map { case(valueName, valueType) =>
      """
        |def with%s(%s: %s) =
        | copy(%s = Option(%s))
      """.stripMargin.format(valueName.capitalize, valueName, valueType, valueName, valueName)}.mkString("") +
    "}"
  }

}
