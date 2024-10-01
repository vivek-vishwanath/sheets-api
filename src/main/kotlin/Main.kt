package org.example

fun main() {
    val spreadsheet = Workbook("18S5ILfkgQgObPI7doq-NK0-5yhHjAWamttyJyy4Y6HE")
    spreadsheet.createNewSheet("NewSheet3")
    spreadsheet.writeSheetData("A1:C2", arrayOf(arrayOf("A1"), arrayOf("A2", "B2")))
    spreadsheet.flush()
}