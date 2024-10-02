package org.example

import org.example.cells.Format

fun main() {
    Workbook("18S5ILfkgQgObPI7doq-NK0-5yhHjAWamttyJyy4Y6HE").apply {
        swapSheet("NewSheet3")
        setFormat("H11", Format(numberFormat = Format.PredefinedFormat.CURRENCY))
    }
}