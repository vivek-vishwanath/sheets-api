package org.example

import cells.Cell
import org.example.cells.Range

fun main() {
    val spreadsheet = Workbook("18S5ILfkgQgObPI7doq-NK0-5yhHjAWamttyJyy4Y6HE")
    println(spreadsheet.readRange("D3:E4"))

    println(Range("D3:E4"))
    println(Range("ZX2:VU102842"))
    println(Cell("AAAAA345678"))
}