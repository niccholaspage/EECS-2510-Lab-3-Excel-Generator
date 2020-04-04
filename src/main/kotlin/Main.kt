import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.BufferedReader
import java.io.File
import java.io.FileOutputStream
import java.io.InputStreamReader
import java.util.concurrent.Executors

const val NUMBER_OF_RUNS = 3

val replacements = mapOf(
    "Balance Factor Changes" to "BF Changes",
    "A to Y Balance Factor Changes" to "A to Y BF Changes"
)

val ignore = listOf(
    "Number of Items",
    "Elapsed Time"
)

val datatypes = listOf(
    "RBT",
    "AVL",
    "BST",
    "Skip List"
)

fun runProgram(executableFile: File, filePath: String): MutableMap<String, MutableMap<String, Any>> {
    val builder = ProcessBuilder()

    // Windows check would go here

    builder.command("\"${executableFile.path}\" \"$filePath\"")

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
                break
            }
        }
    }

    // TODO: Update C++ program to return exit codes properly and utilize this
    val exitCode = process.waitFor()

    executor.shutdown()

    var currentDatatype: String? = null

    val trackingStats = mutableMapOf<String, MutableMap<String, Any>>()

    datatypes.forEach { datatype ->
        trackingStats[datatype] = mutableMapOf<String, Any>("File" to filePath.substringAfterLast('\\'))
    }

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

    return trackingStats
}

fun main(args: Array<String>) {
    if (args.isEmpty()) {
        println(
            """Lab 3 Excel Generator
            |Nicholas Nassar
            |This program runs the structure benchmarking application for Lab 3 on specified files multiple times,
            |averaging out the results and writing them to out to an Excel workbook.
            |
            |The first argument should be the path to the benchmarking application, and every subsequent argument should
            |be the text files you would like to test. If using recursive mode, only the application path needs to be
            |specified.
            |
            |Available Flags:
            |-r - Recursive mode: Any .txt sitting in the working directory directories underneath it will be tested
        """.trimMargin()
        )
        println("No file name specified!")

        return
    }

    val executablePath = args[0]

    val executableFile = File(executablePath)

    if (!executableFile.exists()) {
        println("The benchmarking application was not found at the provided path!")

        return
    }

    val filePaths = if (args[0] == "-r") {
        val workingDirectory = File(".")

        workingDirectory.walk().filter { it.extension == "txt" }.map { it.canonicalPath }.toList()
    } else {
        args.toList().subList(1, args.size)
    }

    if (filePaths.isEmpty()) {
        println("No text files found or supplied via arguments!")

        return
    }

    val workbook = XSSFWorkbook()

    var dataRow = 1

    for (filePath in filePaths) {
        println("Running $filePath:")

        val allStats = mutableListOf<MutableMap<String, MutableMap<String, Any>>>()

        for (i in 0 until NUMBER_OF_RUNS) {
            val stats = runProgram(executableFile, filePath)

            allStats.add(stats)

            println("Finished Test ${i + 1} / $NUMBER_OF_RUNS")
        }

        println()

        val firstStats = allStats[0]

        for (datatype in datatypes) {
            for (i in 0 until NUMBER_OF_RUNS) {
                firstStats[datatype]!!["Time ${i + 1}"] = allStats[i][datatype]!!["Elapsed Time"] as Double
            }
        }

        datatypes.forEach { datatype ->
            val sheet = workbook.getSheet(datatype) ?: workbook.createSheet(datatype)

            val header = sheet.createRow(0)

            val data = sheet.createRow(dataRow)

            var currentColumn = 0

            allStats[0][datatype]!!.forEach writer@{ key, value ->
                if (ignore.contains(key)) {
                    return@writer
                }

                val replacement = replacements[key]

                if (replacement != null) {
                    header.createCell(currentColumn).setCellValue(replacement)
                } else {
                    header.createCell(currentColumn).setCellValue(key)
                }

                when (value) {
                    is String -> {
                        data.createCell(currentColumn).setCellValue(value)
                    }
                    is Double -> {
                        data.createCell(currentColumn).setCellValue(value)
                    }
                    else -> {
                        println("Bad type!!")
                    }
                }

                currentColumn++
            }

            val timeColumns = header.filter { it.stringCellValue.startsWith("Time ") }

            val firstTimeHeader = timeColumns.first()
            val lastTimeHeader = timeColumns.last()

            val firstTimeCell = CellReference(data.getCell(firstTimeHeader.columnIndex)).formatAsString(false)
            val lastTimeCell = CellReference(data.getCell(lastTimeHeader.columnIndex)).formatAsString(false)

            header.createCell(currentColumn).setCellValue("Average Time")
            data.createCell(currentColumn).cellFormula = "AVERAGE($firstTimeCell:$lastTimeCell)"
        }

        dataRow++
    }

    workbook.write(FileOutputStream("output.xlsx"))

    workbook.close()
}