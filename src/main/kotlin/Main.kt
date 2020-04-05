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
 * my jobs. Kotlin greatly cuts down on the verbosity seen in Java and in my opinion, is actually fun to write.
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
    // in our stats. Finally, we call the toMap method, which is an extension method that
    // converts an iterable of pairs into a map.
    val trackingStats = datatypes.map {
        it to mutableMapOf<String, Any>()
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

/**
 * The entry point into the program. We process the commandline arguments here
 * to determine what executable file we will be running and what text files
 * tests will actually be run on.
 */
fun main(args: Array<String>) {
    if (args.isEmpty()) { // If the user didn't provide any arguments, we don't know what to do!
        // Print out a helpful message explaining what our program does and how to use it.
        // We utilize Kotlin's raw string feature to make the string look better versus
        // having to use string concatenation and newline characters. We call trimMargin
        // on our string, which will return another string without the whitespace and | character.
        println(
            """Lab 3 Excel Generator
            |Nicholas Nassar
            |This program runs the structure benchmarking application for Lab 3 on specified files multiple times,
            |averaging out the results and writing them to out to an Excel workbook titled output.xlsx alongside
            |the current working directory of this program.
            |
            |The first argument should be the path to the benchmarking application, and every subsequent argument should
            |be the text files you would like to test. If using recursive mode, only the application path needs to be
            |specified.
            |
            |Available Flags:
            |-r - Recursive mode: Any .txt sitting in the working directory directories underneath it will be tested
        """.trimMargin()
        )

        return // We have nothing else to do, so we return
    }

    // The first argument should be the executable, but we need to exclude any flags the user
    // ran the program with, so we filter out any arguments that start with a dash, then grab
    // the first one from the filtered list, or null if it doesn't exist.
    val executablePath = args.filterNot { it.startsWith("-") }.firstOrNull()

    if (executablePath == null) { // The user didn't provide an executable path,
        println("No benchmarking application specified!") // so we print out a message since we don't know what to run!

        return // We have nothing else to do.
    }

    val executableFile = File(executablePath) // Construct a File object so we can check if the executable exists.

    if (!executableFile.exists()) { // If the executable file doesn't exist, we can't run it!
        println("The benchmarking application was not found at the provided path!") // Tell the user the file does not exist.

        return // We're done.
    }

    // This variable will store all paths to all of the text files we want to run
    // our benchmark program on. Depending on whether the program was launched with
    // the -r flag or not, we either have to recursively walk through the working
    // directory, or just take the arguments as the file paths.
    val filePaths = if (args[0] == "-r") { // If the first argument is the -r flag...
        val workingDirectory = File(".") // We get the working directory by passing . into the File constructor

        // We walk through our working directory, filtering by the .txt extension so we only get text files, map
        // each file to its canonical path, which will give us the full path of each file, and finally turn it into
        // a list.
        workingDirectory.walk().filter { it.extension == "txt" }.map { it.canonicalPath }.toList()
    } else { // We didn't get the -r flag, so we can just run files based on the arguments
        // We just need to get the text file paths from the argument list, so we convert our
        // args array into a list, then take a sublist of that list, excluding the first item,
        // since the first item refers to the executable file path.
        args.toList().subList(1, args.size)
    }

    // If our filePaths list is empty, the user didn't supply any text files or none were found
    // when recursively walking through our working directory.
    if (filePaths.isEmpty()) {
        println("No text files found or supplied via arguments!") // Tell the user the issue.

        return // Return since we have no text files to run tests on.
    }

    val workbook = XSSFWorkbook() // Construct a new .xlsx Excel workbook.

    // We will start writing data at row 2 of the Excel sheet.
    // Apache POI is zero-based, so we start at 1.
    var currentRow = 1

    for (filePath in filePaths) {       // Loop through each file path,
        println("Running $filePath:")   // and tell the user which file is being run.

        // We run our benchmarking program on the file n number of times,
        // n being our NUMBER_OF_RUNS constant. To do this, we map an int
        // range from 0 to NUMBER_OF_RUNS to the stats from a run of the program.
        val runData = (0 until NUMBER_OF_RUNS).map { i ->
            val stats = runProgram(executableFile, filePath) // Run the program and get the results

            println("Finished Test ${i + 1} / $NUMBER_OF_RUNS") // Tell the user we just finished running a test

            // Return the stats to our map function. Since we aren't using
            // the return keyword, the last statement, which is the line below,
            // will implicitly be returned to our map call.
            stats
        }

        println() // Print a newline just for formatting, as we've just finished running all tests for a file.

        // We build a map that will be used to sum all of the stats from each test result.
        // To do this, we map each datatype to a key value pair, with the datatype being the key,
        // and a mutable map of strings and any type as the value. We include a key value pair for
        // File, so that the file name eventually gets written out to our Excel sheet. After mapping
        // out our datatypes, we call toMap to convert our pairs to an actual map.
        val sumOfAllStats = datatypes.map {
            it to mutableMapOf<String, Any>("File" to filePath.substringAfterLast('\\'))
        }.toMap()

        runData.forEach { stats -> // Loop through each of our runs,
            stats.forEach { (datatype, results) -> // Loop through the stats for each datatype,
                val sumResults =
                    sumOfAllStats[datatype]!! // Get the map that will be storing the sum of all of our stats

                // Loop through the results for our datatype so we can look at each stat
                results.forEach { (statName, statValue) ->
                    when (statValue) { // When a stat value is...
                        is Double -> { // a double,
                            // We can get the current sum from sumResults, or default to 0 if there
                            // isn't one yet.
                            val sumValue = sumResults.getOrDefault(statName, 0.0) as Double

                            // We then add our statValue to the sumValue,
                            // and place it back into the map.
                            sumResults[statName] = sumValue + statValue
                        }
                        else -> { // If we have some other type here, something weird has happened.
                            throw IllegalArgumentException("Not seen type!") // Throw an exception!
                        }
                    }
                }
            }
        }

        // At this point, we have the sum of all of our stats, so we can write to our Excel workbook.
        datatypes.forEach { datatype -> // Loop through each of our datatypes.
            // We will get the sheet for our datatype, creating one if it doesn't already exist.
            val sheet = workbook.getSheet(datatype) ?: workbook.createSheet(datatype)

            // NOTE: THIS HEADER GENERATION CODE GETS CALLED EVERY RUN.
            // This could be optimized so that the header only gets written once.
            val headerRow = sheet.createRow(0) // We create a header row for the sheet.

            val dataRow = sheet.createRow(currentRow) // We also create a data row so we can write the data for our file.

            var currentColumn = 0 // We start writing at column 0.

            // Get the sum of all stats for the datatype, and use not-null assertion
            // because we know that our datatype is in the map. This can probably
            // be written in a nicer way.
            // We label this for each loop writer, so we can return to it later.
            sumOfAllStats[datatype]!!.forEach writer@{ (key, value) -> // Loop through each stat,
                if (ignore.contains(key)) { // If we are ignoring this key,
                    return@writer // we return to our writer forEach loop. This is basically like a continue.
                }

                // If we have a replacement for this key, it will be
                // the header value. Otherwise, we will just use the
                // key as the header value. This allows us to modify
                // header names so that they are different from the
                // benchmark program's output.
                val headerValue = replacements[key] ?: key

                // Create the cell for our header value in our header row
                // at the current column, and set its value to the header
                // value.
                headerRow.createCell(currentColumn).setCellValue(headerValue)

                val dataCell = dataRow.createCell(currentColumn)

                when (value) { // When the value is...
                    is String -> { // a string,
                        dataCell.setCellValue(value) // we set the value of the cell to the value.
                    }
                    is Double -> { // a double,
                        // we first average out the value by dividing by the number of runs,
                        val averageValue = value / NUMBER_OF_RUNS

                        dataCell.setCellValue(averageValue) // and set the cell value to our average value.
                    }
                    else -> { // If we have some other type here, something weird has happened.
                        throw IllegalArgumentException("Not seen type!") // Throw an exception!
                    }
                }

                currentColumn++ // Increment our current column, as we are moving onto the next one
            }
        }

        currentRow++ // Increment our current row, since we just finished the row we are currently on.
    }

    workbook.write(FileOutputStream("output.xlsx")) // Write our workbook output to output.xlsx

    workbook.close() // Close out the workbook. We are finished!
}