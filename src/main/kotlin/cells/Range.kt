package org.example.cells

import cells.Cell

class Range(col: Int, row: Int, val width: Int, val height: Int): Cell(col, row) {

    constructor(range: Range) : this(range.col, range.row, range.width, range.height)

    constructor(address: String) : this(fromString(address))

    companion object {

        fun fromString(address: String): Range {
            val parts = address.split(":")
            if (parts.size == 1) throw InvalidAddressException("Range address is missing delimiter `:`")
            if (parts.size > 2) throw InvalidAddressException("Range address has too many delimiters `:`")
            val topLeft = Cell(parts[0])
            val bottomRight = Cell(parts[1])
            return Range(topLeft.col, topLeft.row, bottomRight.col - topLeft.col + 1, bottomRight.row - topLeft.row + 1)
        }
    }
}