package org.example

import org.example.cells.Format

fun main() {
    val spreadsheet = Workbook("18S5ILfkgQgObPI7doq-NK0-5yhHjAWamttyJyy4Y6HE")
    spreadsheet.writeSheetData("C23:C27", arrayOf(arrayOf("Hello", "World", "A"), arrayOf("hey", "D", "E")))
    spreadsheet.flush()
}