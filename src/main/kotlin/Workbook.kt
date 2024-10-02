package org.example

import com.google.api.services.sheets.v4.model.*
import org.example.OAuth.service
import org.example.cells.Format
import org.example.cells.MismatchedDimensionsException
import org.example.cells.Range
import java.lang.IllegalArgumentException

class Workbook(val spreadsheetID: String) {

    private fun fetch() = service.spreadsheets().get(spreadsheetID).execute().sheets

    val sheets by lazy { fetch() }

    var sheet = sheets[0]

    val requests = mutableListOf<Request>()

    fun createNewSheet(newSheetName: String) {
        requests.add(Request().apply {
            addSheet = AddSheetRequest().apply {
                properties = SheetProperties().apply {
                    title = newSheetName
                }
            }
        })
        flush()
        swapSheet(newSheetName)
    }

    fun swapSheet(name: String) {
        sheets.forEach {
            if (it.properties.title == name) {
                sheet = it
                return
            }
        }
        throw IllegalArgumentException("Invalid sheet name: \"$name\"")
    }

    fun readRange(address: String): Array<Array<String>> {
        flush()
        val response: ValueRange = service.spreadsheets().values()
            .get(spreadsheetID, address)
            .execute()
        return response.getValues()?.map {
            it.map { that -> that.toString() }.toTypedArray()
        }?.toTypedArray() ?: emptyArray<Array<String>>()
    }

    fun writeSheetData(address: String, values: Array<Array<String>>) {
        val targetRange = Range(address)
        if (values.size > targetRange.height + 1)
            throw MismatchedDimensionsException()
        else if (values.size <= targetRange.height)
            println("WARNING: Mismatched Dimensions between ${values.size} rows from address and ${targetRange.height} rows of data")
        requests.add(
            Request().setUpdateCells(UpdateCellsRequest().apply {
                range = targetRange.gridRange(sheet.properties.sheetId)
                rows = values.map {
                    if (it.size > targetRange.width + 1)
                        throw MismatchedDimensionsException()
                    else if (it.size <= targetRange.width)
                        println("WARNING: Mismatched Dimensions between ${it.size} columns from address and ${targetRange.width} columns of data")
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

    fun addCheckbox(address: String) {
        requests.add(Request().apply {
            setDataValidation = SetDataValidationRequest().apply {
                range = Range(address).gridRange(sheet.properties.sheetId)
                rule = DataValidationRule().apply {
                    condition = BooleanCondition().apply {
                        type = "BOOLEAN"
                    }
                    strict = false
                    showCustomUi = true
                }
            }
        })
    }

    fun mergeCells(address: String) {
        requests.add(Request().apply {
            mergeCells = MergeCellsRequest().apply {
                range = Range(address).gridRange(sheet.properties.sheetId)
                mergeType = "MERGE_ALL"
            }
        })
    }

    fun resize(intRange: IntRange, rows: Boolean, newWidth: Int? = null) {
        val dimensionRange = DimensionRange().apply {
            sheetId = sheet.properties.sheetId
            dimension = if (rows) "ROWS" else "COLUMNS"
            startIndex = intRange.first - 1
            endIndex = intRange.last
        }
        requests.add(Request().apply {
            if (newWidth == null)
                autoResizeDimensions = AutoResizeDimensionsRequest().setDimensions(dimensionRange)
            else
                updateDimensionProperties = UpdateDimensionPropertiesRequest().apply {
                    range = dimensionRange
                    properties = DimensionProperties().apply {
                        pixelSize = newWidth
                    }
                    fields = "pixelSize"
                }
        })
    }

    fun protectRange(address: String, editors: Array<String>) {
        requests.add(Request().apply {
            addProtectedRange = AddProtectedRangeRequest().apply {
                protectedRange = ProtectedRange().apply {
                    this.range = Range(address).gridRange(sheet.properties.sheetId)
                    this.description = "Protected range example"
                    this.editors = Editors().setUsers(editors.toList())
                    this.warningOnly = false
                }
            }
        })
    }

    fun flush() {
        if (requests.size == 0) return
        val batchUpdateRequest = BatchUpdateSpreadsheetRequest().setRequests(requests)
        service.spreadsheets().batchUpdate(spreadsheetID, batchUpdateRequest).execute()
        requests.clear()
        sheets.clear()
        sheets.addAll(fetch())
    }
}