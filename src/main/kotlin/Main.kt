import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream
import java.util.*

val replacements = mapOf(
    "Balance Factor Changes" to "BF Changes",
    "A to Y Balance Factor Changes" to "A to Y BF Changes"
)

val ignore = listOf(
    "Number of Items"
)

val datatypes = listOf(
    Datatype("RBT"),
    Datatype("AVL"),
    Datatype("BST"),
    Datatype("Skip List")
)

fun main() {
    val scanner = Scanner(System.`in`)

    val input = mutableListOf<String>()

    println("Paste your program output:")

    while (scanner.hasNextLine()) {
        val line = scanner.nextLine()

        if (line.isEmpty()) {
            break;
        }

        input.add(line)
    }

    var currentDatatype: Datatype? = null

    val trackingStats = mutableMapOf<Datatype, MutableMap<String, Double>>()

    datatypes.forEach { datatype ->
        trackingStats[datatype] = mutableMapOf()
    }

    val fileName = input[0]

    input.subList(1, input.size).forEach { line ->
        datatypes.forEach typeCheck@{ datatype ->
            if (line.contains(datatype.name)) {
                currentDatatype = datatype

                return@forEach
            }
        }

        if (line.contains(':')) {
            val key = line.substringBefore(':')

            val count = line.substringAfter(": ").substringBefore(' ')

            trackingStats[currentDatatype]!![key] = count.toDouble()
        }
    }

    val workbook = XSSFWorkbook()

    datatypes.forEach { datatype ->
        val sheet = workbook.createSheet(datatype.name)

        val header = sheet.createRow(0)

        val data = sheet.createRow(1)

        header.createCell(0).setCellValue("File")
        data.createCell(0).setCellValue(fileName.substringAfterLast("\\"))

        var currentColumn = 1

        trackingStats[datatype]!!.forEach writer@ { key, count ->
            if (ignore.contains(key)) {
                return@writer
            }

            val replacement = replacements[key]

            if (replacement != null) {
                header.createCell(currentColumn).setCellValue(replacement)
            } else {
                header.createCell(currentColumn).setCellValue(key)
            }

            data.createCell(currentColumn).setCellValue(count)

            currentColumn++
        }
    }

    workbook.write(FileOutputStream("output.xlsx"))

    workbook.close()
}