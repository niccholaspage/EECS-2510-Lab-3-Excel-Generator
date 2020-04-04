/**
 * Main.kt - Lab 3 Excel Generator
 *
 * This program runs the comparison program written in C++ for Lab 3. It does so by launching
 * the comparison program on files passed in as parameters multiple times, averaging out the results
 * from the program (parsed through the console output of the program) and writing them directly into
 * an Excel workbook. This program utilizes the Kotlin standard library as well as Apache POI and Apache
 * POI OOXML, allowing it to create and manipulate Excel workbooks.
 *
 * Author:     Nicholas Nassar, University of Toledo
 * Class:      EECS 2510-001 Non-Linear Data Structures, Spring 2020
 * Instructor: Dr.Thomas
 * Date:       Apr 4, 2020
 * Copyright:  Copyright 2020 by Nicholas Nassar. All rights reserved.
 */
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.BufferedReader
import java.io.File
import java.io.FileOutputStream
import java.io.InputStreamReader

const val NUMBER_OF_RUNS = 10 // Here, we specify the number of times we will execute our comparison program.

val replacements = mapOf(
    "Balance Factor Changes" to "BF Changes",
    "A to Y Balance Factor Changes" to "A to Y BF Changes"
)

// Stats that don't need to be written out to the Excel file can be placed here so that they are ignored.
// Time is ignored as it is written out for each run instead of a single elapsed time.
val ignore = listOf(
    "Number of Items",
    "Elapsed Time"
)

// The list of datatypes that we will be looking for in the comparison program's output.
val datatypes = listOf(
    "RBT",
    "AVL",
    "BST",
    "Skip List"
)

/**
 * This method runs the program at the given executable file. The given file path will be run through
 * the comparison program. Each line of the program's output will be parsed into a map, with keys for
 * each datatype and values being another map representing the keys and values of each stat.
 */
fun runProgram(executableFile: File, filePath: String): Map<String, MutableMap<String, Any>> {
    // Create a ProcessBuilder, passing in the path of our executable file and the text file we will be running
    // our comparison program on. Paths are enclosed in double quotes to account for any spaces in these file
    // paths.
    val builder = ProcessBuilder("\"${executableFile.path}\" \"$filePath\"")

    val process = builder.start() // Start the process!

    val inputStream = process.inputStream // Retrieve the process's input stream, allowing us to read its console output

    val reader =
        BufferedReader(InputStreamReader(inputStream)) // We use a buffered reader so we can read output line by line

    val output = mutableListOf<String>() // Construct a list we will add each line of output to

    // In Kotlin, assignments are not allowed in expressions! This stops us from doing the following code allowed in Java:
    // String line;
    // while ((line = reader.readLine()) != null) {}
    // So instead, we start off our line variable with the first line from the reader.
    var line = reader.readLine()

    while (true) {
        output.add(line) // We add our line to the output list.

        line = reader.readLine() // Read a new line

        // Our line is null, so we have no more output, so we break out of the while loop!
        if (line == null) {
            break
        }
    }

    // TODO: Update C++ program to return exit codes properly and utilize this
    val exitCode = process.waitFor()

    process.destroy() // Just for cleanup, we destroy the process

    var currentDatatype: String? = null

    val trackingStats =
        datatypes.map { it to mutableMapOf<String, Any>("File" to filePath.substringAfterLast('\\')) }.toMap()

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

        val allStats = mutableListOf<Map<String, MutableMap<String, Any>>>()

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