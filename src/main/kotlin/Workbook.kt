package org.example

import com.google.api.services.sheets.v4.model.ValueRange
import org.example.OAuth.service

class Workbook(val spreadsheetID: String) {

    val sheets by lazy { service.spreadsheets().get(spreadsheetID).execute().sheets }

    var sheet = sheets[0]

    fun readRange(address: String): List<List<String>> {
        val response: ValueRange = service.spreadsheets().values()
            .get(spreadsheetID, address)
            .execute()
        return response.getValues()?.map { it.map { that -> that.toString() } } ?: emptyList()
    }

    fun writeSheetData(address: String, values: List<List<Any>>) {
        val body = ValueRange().setValues(values)
        service.spreadsheets().values()
            .update(spreadsheetID, address, body)
            .setValueInputOption("RAW")
            .execute()
    }
}