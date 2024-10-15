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
    val numberFormat: NumberFormat = Automatic,
    val wrapStrategy: WrapStrategy = WrapStrategy.WRAP,
    val textRotation: TextRotation = AngledText(0),
    val borders: Edges = Edges(),
    val padding: Padding = Padding(),
    val hyperlink: Boolean = false
) {

    data class Color(val r: UByte = 0u, val g: UByte = r, val b: UByte = r) {

        companion object {
            val BLACK = Color()
            val LT_GRAY = Color(225u)
            val YELLOW = Color(0xFFu, 0xFFu, 0u)
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

    sealed interface TextRotation
    class AngledText(val angle: Int = 0): TextRotation
    data object VerticalText: TextRotation

    data class Padding(val all: Int = 0, val top: Int = all, val right: Int = all, val bottom: Int = all, val left: Int = all) {
        fun convert() = com.google.api.services.sheets.v4.model.Padding().apply {
            top = this@Padding.top
            right = this@Padding.right
            bottom = this@Padding.bottom
            left = this@Padding.left
        }
    }

    sealed interface NumberFormat
    enum class PredefinedFormat: NumberFormat { NUMBER, CURRENCY, PERCENT, DATE, TIME, TEXT, SCIENTIFIC }
    data object Automatic: NumberFormat
    data class NumberPattern(val pattern: String): NumberFormat

    class Edges(val all: Edge = Edge(color = Color.LT_GRAY), val top: Edge = all, val right: Edge = all, val bottom: Edge = all, val left: Edge = all) {
        fun convert() = Borders().apply {
            top = this@Edges.top.convert()
            right = this@Edges.right.convert()
            bottom = this@Edges.bottom.convert()
            left = this@Edges.left.convert()
        }

        companion object {
            val ALL = Edges(Edge())
        }
    }

    class Edge(val color: Color = Color.BLACK, val style: EdgeStyle = EdgeStyle.SOLID) {

        fun convert() = Border().apply {
            color = this@Edge.color.convert()
            style = this@Edge.style.toString()
        }
    }

    enum class EdgeStyle {
        NONE, SOLID, DASHED, DOTTED, SOLID_MEDIUM, SOLID_THICK, DOUBLE
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
        numberFormat = NumberFormat().apply {
            when (val nf = this@Format.numberFormat) {
                is NumberPattern -> pattern = nf.pattern
                is PredefinedFormat -> type = nf.toString()
                Automatic -> {}
            }
        }
        textRotation = TextRotation().apply {
            when (val tr = this@Format.textRotation) {
                is AngledText -> angle = tr.angle
                is VerticalText -> vertical = true
            }
        }
        hyperlinkDisplayType = if (hyperlink) "LINKED" else "PLAIN_TEXT"
        padding = this@Format.padding.convert()
    }
}