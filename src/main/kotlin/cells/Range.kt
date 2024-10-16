package org.example.cells

import cells.Cell
import com.google.api.services.sheets.v4.model.GridRange

class Range(col: Int, row: Int, val width: Int, val height: Int): Cell(col, row) {

    constructor(range: Range) : this(range.col, range.row, range.width, range.height)

    constructor(address: String) : this(fromString(address))

    companion object {

        fun fromString(address: String): Range {
            val parts = address.split(":")
            if (parts.size > 2) throw InvalidAddressException("Range address has too many delimiters `:`")
            val topLeft = Cell(parts[0])
            val bottomRight = Cell(if (parts.size == 2) parts[1] else parts[0])
            return Range(topLeft.col, topLeft.row, bottomRight.col - topLeft.col, bottomRight.row - topLeft.row)
        }
    }
    
    fun gridRange(sheetID: Int) = GridRange().apply {
        sheetId = sheetID
        startRowIndex = row - 1
        endRowIndex = row + height
        startColumnIndex = col - 1
        endColumnIndex = col + width
    }

    override operator fun plus(cell: Cell) = Range(col + cell.col, row + cell.row, width, height)

    override fun toString() = "${super.toString()}:${Cell(col + width, row + height)}"
}