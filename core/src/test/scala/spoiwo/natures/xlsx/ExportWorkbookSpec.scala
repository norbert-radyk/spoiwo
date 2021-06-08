package spoiwo.natures.xlsx

import spoiwo.model._
import java.io.ByteArrayOutputStream
import Model2XlsxConversions._
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class ExportWorkbookSpec extends AnyFlatSpec with Matchers {

  private val workbook = Workbook(Sheet(name = "Test Sheet", rows = List.empty))

  "Writing workbook to an output stream" should "write non empty content" in {
    val baos = new ByteArrayOutputStream()
    workbook.writeToOutputStream(baos)
    baos.size() should be > 0
  }

}
