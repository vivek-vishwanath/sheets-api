package org.example

fun main() {
    val spreadsheet = Workbook("18S5ILfkgQgObPI7doq-NK0-5yhHjAWamttyJyy4Y6HE")
    spreadsheet.swapSheet("NewSheet3")
    spreadsheet.resize(3..5, false)
    spreadsheet.flush()
}