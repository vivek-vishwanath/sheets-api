package org.example

import com.google.api.services.sheets.v4.model.*
import org.example.OAuth.service
import org.example.cells.Format
import org.example.cells.Range

class Workbook(val spreadsheetID: String) {

    val sheets by lazy { service.spreadsheets().get(spreadsheetID).execute().sheets }

    var sheet = sheets[0]

    val requests = mutableListOf<Request>()

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

    fun setFormat(address: String, format: Format) {
        val targetRange = Range(address)
        requests.add(
            Request().apply {
                updateCells = UpdateCellsRequest().apply {
                    rows = Array(targetRange.height + 1) {
                        RowData().setValues(Array(targetRange.width + 1){
                            CellData().setUserEnteredFormat(format.convert())
                        }.toList()) }.toList()
                    range = GridRange().apply {
                        sheetId = sheet.properties.sheetId
                        startRowIndex = targetRange.row - 1
                        endRowIndex = targetRange.row + targetRange.height
                        startColumnIndex = targetRange.col - 1
                        endColumnIndex = targetRange.col + targetRange.width
                    }
                    fields = "userEnteredFormat"
                }
            }
        )
    }

    fun flush() {
        val batchUpdateRequest = BatchUpdateSpreadsheetRequest().setRequests(requests)
        service.spreadsheets().batchUpdate(spreadsheetID, batchUpdateRequest).execute()
        requests.clear()
    }
}