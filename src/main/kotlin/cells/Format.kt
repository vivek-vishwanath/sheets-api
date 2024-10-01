package org.example.cells

import com.google.api.services.sheets.v4.model.*

data class Format(
    val bold: Boolean = false,
    val italic: Boolean = false,
    val underline: Boolean = false,
    val strikethrough: Boolean = false,
    val fontFamily: String = "Calibri",
    val fontSize: Int = 12,
    val foreground: Color = Color(0u, 0u, 0u),
    val background: Color = Color(0xFFu, 0xFFu, 0xFFu),
    val horizontalAlignment: HorizontalAlignment = HorizontalAlignment.CENTER,
    val verticalAlignment: VerticalAlignment = VerticalAlignment.MIDDLE,
    val numberFormat: NumberFormat = NumberFormat(),
    val wrapStrategy: WrapStrategy = WrapStrategy.WRAP,
    val borders: Edges = Edges()
) {

    data class Color(val r: UByte = 0u, val g: UByte = 0u, val b: UByte = 0u) {

        companion object {
            val BLACK = Color()
        }
        fun convert() = com.google.api.services.sheets.v4.model.Color().apply {
            red = r.toFloat() / 256
            green = g.toFloat() / 256
            blue = b.toFloat() / 256
        }
    }

    enum class HorizontalAlignment {
        LEFT, CENTER, RIGHT, JUSTIFIED
    }

    enum class VerticalAlignment {
        TOP, MIDDLE, BOTTOM, JUSTIFIED, BASELINE
    }

    enum class WrapStrategy {
        OVERFLOW_CELL, LEGACY_WRAP, CLIP, WRAP
    }

    class Edges(val top: Edge = Edge(), val right: Edge = Edge(), val bottom: Edge = Edge(), val left: Edge = Edge()) {
        fun convert() = Borders().apply {
            top = this@Edges.top.convert()
            right = this@Edges.right.convert()
            bottom = this@Edges.bottom.convert()
            left = this@Edges.left.convert()
        }
    }

    class Edge(val color: Color = Color.BLACK, val style: EdgeStyle = EdgeStyle.NONE) {

        fun convert() = Border().apply {
            color = this@Edge.color.convert()
            style = this@Edge.style.toString()
            width = 4
        }
    }

    enum class EdgeStyle {
        NONE, SOLID, DOTTED, DASHED, DOUBLE, GROOVE, RIDGE, INSET, OUTSET,
    }

    fun convert() = CellFormat().apply {
        backgroundColor = this@Format.background.convert()
        verticalAlignment = this@Format.verticalAlignment.toString()
        horizontalAlignment = this@Format.horizontalAlignment.toString()
        wrapStrategy = this@Format.wrapStrategy.toString()
        textFormat = TextFormat().apply {
            fontSize = this@Format.fontSize
            fontFamily = this@Format.fontFamily
            bold = this@Format.bold
            italic = this@Format.italic
            underline = this@Format.underline
            strikethrough = this@Format.strikethrough
        }
        borders = this@Format.borders.convert()
    }
}