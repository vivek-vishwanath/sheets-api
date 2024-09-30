package org.example

import com.google.api.services.sheets.v4.model.ValueRange
import org.example.OAuth.service

class Workbook(val spreadsheetID: String) {

    val sheets by lazy { service.spreadsheets().get(spreadsheetID).execute().sheets }

    var sheet = sheets[0]

    fun readCell(col: Int, row: Int): List<List<String>> {
        val response: ValueRange = service.spreadsheets().values()
            .get(spreadsheetID, "${(col + 65).toChar()}${row + 1}")
            .execute()
        return response.getValues()?.map { it.map { that -> that.toString() } } ?: emptyList()
    }
}