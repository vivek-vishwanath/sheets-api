package cells

import org.example.cells.InvalidAddressException

open class Cell(val col: Int, val row: Int) {

    constructor(cell: Cell) : this(cell.col, cell.row)

    constructor(address: String) : this(fromString(address))

    open operator fun plus(cell: Cell) = Cell(col + cell.col, row + cell.row)

    companion object {

        fun fromString(address: String): Cell {
            var i = 0
            var col = 0
            while (address[i] in 'A'..'Z') {
                col *= 26
                col += address[i++] - 'A' + 1
                if (i >= address.length) throw InvalidAddressException()
            }
            val row = address.substring(i).toIntOrNull()
                ?: throw InvalidAddressException("Missing or Invalid Row #: ${address.substring(i)} is not an int")
            return Cell(col, row)
        }
    }

    override fun toString(): String {
        var col = col
        var out = "$row"
        while (col > 0) {
            col--
            out = "${'A' + col % 26}$out"
            col /= 26
        }
        return out
    }
}