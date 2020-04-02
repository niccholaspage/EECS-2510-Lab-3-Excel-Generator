import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.BufferedReader
import java.io.File
import java.io.FileOutputStream
import java.io.InputStreamReader
import java.util.*
import java.util.concurrent.Executors

val replacements = mapOf(
    "Balance Factor Changes" to "BF Changes",
    "A to Y Balance Factor Changes" to "A to Y BF Changes"
)

val ignore = listOf(
    "Number of Items"
)

val datatypes = listOf(
    "RBT",
    "AVL",
    "BST",
    "Skip List"
)

fun runTests(filePath: String): List<String> {
    val builder = ProcessBuilder()

    // Windows check would go here

    val exeFile = File("C:\\Users\\nicch\\Documents\\Projects\\C++\\EECS-2510-Lab-3\\x64\\Release\\Lab 3.exe")

    builder.command("\"${exeFile.path}\" \"$filePath\"")

    val process = builder.start()

    val inputStream = process.inputStream

    val reader = BufferedReader(InputStreamReader(inputStream))

    val executor = Executors.newSingleThreadExecutor()

    val output = mutableListOf<String>()

    executor.submit {
        var line = reader.readLine()

        while (true) {
            output.add(line)

            line = reader.readLine()

            if (line == null) {
                println("BREAK")
                break
            }
        }
    }

    val exitCode = process.waitFor()

    executor.shutdown()

    println("Exit Code: $exitCode")

    return output
}

fun main(args: Array<String>) {
    if (args.isEmpty()) {
        println("No file name specified!")

        return
    }

    val filePath = args[0]

    val scanner = Scanner(System.`in`)

    val output = runTests(filePath)

    var currentDatatype: String? = null

    val trackingStats = mutableMapOf<String, MutableMap<String, Double>>()

    datatypes.forEach { datatype ->
        trackingStats[datatype] = mutableMapOf()
    }

    val fileName = output[0]

    output.subList(1, output.size).forEach { line ->
        datatypes.forEach typeCheck@{ datatype ->
            if (line.contains(datatype)) {
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
        val sheet = workbook.createSheet(datatype)

        val header = sheet.createRow(0)

        val data = sheet.createRow(1)

        header.createCell(0).setCellValue("File")
        data.createCell(0).setCellValue(fileName.substringAfterLast("\\"))

        var currentColumn = 1

        trackingStats[datatype]!!.forEach writer@{ key, count ->
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