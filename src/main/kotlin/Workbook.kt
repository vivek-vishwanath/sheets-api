package org.example

import com.google.api.services.sheets.v4.model.*
import org.example.OAuth.service
import org.example.cells.Format
import org.example.cells.MismatchedDimensionsException
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

    fun writeSheetData(address: String, values: Array<Array<String>>) {
        val targetRange = Range(address)
        if (values.size > targetRange.height + 1)
            throw MismatchedDimensionsException()
        requests.add(
            Request().setUpdateCells(UpdateCellsRequest().apply {
                range = targetRange.gridRange(sheet.properties.sheetId)
                rows = values.map {
                    if (it.size > targetRange.width + 1)
                        throw MismatchedDimensionsException()
                    RowData().setValues(it.map { that ->
                        CellData().setUserEnteredValue(ExtendedValue().apply {
                            if (that.startsWith("=")) formulaValue = that
                            else stringValue = that
                        })
                    })
                }.toMutableList()
                fields = "userEnteredValue"
            })
        )
    }

    fun setFormat(address: String, format: Format) {
        val targetRange = Range(address)
        requests.add(
            Request().apply {
                updateCells = UpdateCellsRequest().apply {
                    rows = Array(targetRange.height + 1) {
                        RowData().setValues(Array(targetRange.width + 1) {
                            CellData().setUserEnteredFormat(format.convert())
                        }.toList())
                    }.toList()
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

    fun add(startRowIndex: Int, numberOfRows: Int, rows: Boolean) {
        requests.add(Request().apply {
            insertDimension = InsertDimensionRequest().apply {
                range = DimensionRange().apply {
                    sheetId = this@Workbook.sheet.properties.sheetId
                    dimension = if (rows) "ROWS" else "COLUMNS"
                    startIndex = startRowIndex
                    endIndex = startRowIndex + numberOfRows
                }
                inheritFromBefore = false
            }
        })
    }

    fun delete(intRange: IntRange, rows: Boolean) {
        requests.add(Request().apply {
            deleteDimension = DeleteDimensionRequest().apply {
                range = DimensionRange().apply {
                    sheetId = this@Workbook.sheet.properties.sheetId
                    dimension = if (rows) "ROWS" else "COLUMNS"
                    startIndex = intRange.first - 1
                    endIndex = intRange.last
                }
            }
        })
    }

    fun flush() {
        val batchUpdateRequest = BatchUpdateSpreadsheetRequest().setRequests(requests)
        service.spreadsheets().batchUpdate(spreadsheetID, batchUpdateRequest).execute()
        requests.clear()
    }
}