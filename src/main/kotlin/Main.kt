package org.example

import org.example.cells.Format

fun main() {
    val spreadsheet = Workbook("18S5ILfkgQgObPI7doq-NK0-5yhHjAWamttyJyy4Y6HE")
    spreadsheet.swapSheet("NewSheet3")
    spreadsheet.setFormat("H10", Format(numberFormat = Format.PredefinedFormat.CURRENCY))
    spreadsheet.setFormat("H12", Format(numberFormat = Format.PredefinedFormat.PERCENT))
    spreadsheet.setFormat("H11", Format(textRotation = Format.VerticalText))
    spreadsheet.flush()
}