package org.example

import org.example.cells.Format

fun main() {
    val spreadsheet = Workbook("18S5ILfkgQgObPI7doq-NK0-5yhHjAWamttyJyy4Y6HE")
    spreadsheet.delete(15..17, true)
    spreadsheet.flush()
}