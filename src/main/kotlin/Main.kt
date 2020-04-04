/**
 * Main.kt - Lab 3 Excel Generator
 *
 * This program runs the comparison program written in C++ for Lab 3. It does so by launching
 * the comparison program on files passed in as parameters multiple times, averaging out the results
 * from the program (parsed through the console output of the program) and writing them directly into
 * an Excel workbook. This program utilizes the Kotlin standard library as well as Apache POI and Apache
 * POI OOXML, allowing it to create and manipulate Excel workbooks.
 *
 * Kotlin and Apache POI were picked due to my experience with the language and dependency through one of
 * my jobs. Kotlin greatly cuts down the verbosity seen in Java and in my opinion, is actually fun to write.
 * Apache POI is a very helpful library and makes creating and editing Excel workbooks straightforward and easy.
 *
 * Author:     Nicholas Nassar, University of Toledo
 * Class:      EECS 2510-001 Non-Linear Data Structures, Spring 2020
 * Instructor: Dr.Thomas
 * Date:       Apr 4, 2020
 * Copyright:  Copyright 2020 by Nicholas Nassar. All rights reserved.
 */
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
val ignore = listOf(
    "Number of Items"
)

// The list of datatypes that we will be looking for in the comparison program's output.
val datatypes = listOf(
    "RBT",
    "AVL",
    "BST",
    "Skip List"
)

// We use a type alias to make this typing a lot less insane to understand.
// A stat map is a map with strings as the key, representing the datatype.
// The value is another map, with the key representing the name of the stat
// we are recording, and the value being the result of the stat, which could
// be a count or time in seconds for example.
typealias StatMap = Map<String, MutableMap<String, Any>>

/**
 * This method runs the program at the given executable file. The given file path will be run through
 * the comparison program. Each line of the program's output will be parsed into a map, with keys for
 * each datatype and values being another map representing the keys and values of each stat.
 */
fun runProgram(executableFile: File, filePath: String): StatMap {
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
    // This most likely could have been written in more idiomatic Kotlin, but this program
    // was written quickly and I didn't want to figure out a better way to do this.
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

    var currentDatatype: String? = null // We need to keep track of which datatype we are recording stats for

    // Lets create a map that will keep track of our stats for each datatype. We start off by using the
    // map extension function provided by the Kotlin standard library, which allows us to transform an
    // iterable of items with a certain type into different objects. In this case, we map our datatype,
    // which is a string, to a mutable map that represents the key-values. We will use this map to throw
    // in our stats. We include a single item into the map at first, which is the name of the file
    // we are running tests on. Finally, we call the toMap method, which is an extension method that
    // converts an iterable of pairs into a map.
    val trackingStats = datatypes.map {
        it to mutableMapOf<String, Any>("File" to filePath.substringAfterLast('\\'))
    }.toMap()

    // We now need to process the output. Since the first line of the comparison program is just
    // the file the program is being ran on, we skip it by getting a sub list of our output,
    // which will start at the second element.
    output.subList(1, output.size).forEach { line -> // Loop through each line of our sublist.
        // See if any of our datatype names are in the line.
        val newDatatype = datatypes.find { line.contains(it) }

        if (newDatatype != null) { // If new datatype isn't null, it looks like we did find a datatype.
            currentDatatype = newDatatype   // We update our current datatype,

            return@forEach                  // and return to the forEach, to parse the next line.
        }

        if (line.contains(':')) { // If our line contains a colon, we know we have a stat.
            val key = line.substringBefore(':') // The key or name of the stat is everything before the colon,

            // and the value for the stat is everything after the colon and a space and before another space,
            // to fix any issues with units, like elapsed time having the seconds unit.
            val value = line.substringAfter(": ").substringBefore(' ')

            // We now get the map of key values for our current datatype,
            // and set the value at the key to the value we just parsed.
            trackingStats[currentDatatype]!![key] = value.toDouble()
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

        return
    }

    val executablePath = args.filterNot { it.startsWith("-") }.firstOrNull()

    if (executablePath == null) {
        println("No benchmarking application specified!")

        return
    }

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

        val allStats = mutableListOf<StatMap>()

        for (i in 0 until NUMBER_OF_RUNS) {
            val stats = runProgram(executableFile, filePath)

            allStats.add(stats)

            println("Finished Test ${i + 1} / $NUMBER_OF_RUNS")
        }

        println()

        val averageStats = datatypes.map {
            it to mutableMapOf<String, Any>("File" to filePath.substringAfterLast('\\'))
        }.toMap()

        allStats.forEach { stats ->
            stats.forEach { (datatype, results) ->
                val averageResults = averageStats[datatype]!!

                results.forEach { (statName, statValue) ->
                    when (statValue) {
                        is Double -> {
                            val averageValue = averageResults.getOrDefault(statName, 0.0) as Double

                            averageResults[statName] = averageValue + statValue
                        }
                        is String -> {
                            averageResults[statName] = statValue
                        }
                        else -> {
                            throw IllegalArgumentException("Not seen type!")
                        }
                    }
                }
            }
        }

        datatypes.forEach { datatype ->
            val sheet = workbook.getSheet(datatype) ?: workbook.createSheet(datatype)

            val header = sheet.createRow(0)

            val data = sheet.createRow(dataRow)

            var currentColumn = 0

            averageStats[datatype]!!.forEach writer@{ key, value ->
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
                        val averagedValue = value / NUMBER_OF_RUNS

                        data.createCell(currentColumn).setCellValue(averagedValue)
                    }
                    else -> {
                        println("Bad type!!")
                    }
                }

                currentColumn++
            }
        }

        dataRow++
    }

    workbook.write(FileOutputStream("output.xlsx"))

    workbook.close()
}