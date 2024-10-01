package org.example

import org.example.cells.Format

fun main() {
    val spreadsheet = Workbook("18S5ILfkgQgObPI7doq-NK0-5yhHjAWamttyJyy4Y6HE")
    spreadsheet.setFormat("C14:C14",
        Format(
            background = Format.Color(g = 255u, b = 255u),
            verticalAlignment = Format.VerticalAlignment.TOP,
            horizontalAlignment = Format.HorizontalAlignment.RIGHT,
            wrapStrategy = Format.WrapStrategy.CLIP,
            bold = true,
            strikethrough = true,
            italic = true,
            underline = true,
            fontFamily = "Trebuchet MS",
            fontSize = 30,
            borders = Format.Edges(
                left = Format.Edge(
                    color = Format.Color(r=200u, g=100u),
                    style = Format.EdgeStyle.SOLID,
                )
            )
        ))
    spreadsheet.flush()
}