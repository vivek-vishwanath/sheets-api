package org.example

import cells.Cell
import com.google.api.services.sheets.v4.model.*
import org.example.OAuth.service
import org.example.cells.Format
import org.example.cells.MismatchedDimensionsException
import org.example.cells.Range
import java.lang.IllegalArgumentException

class Workbook(val spreadsheetID: String) {

    private fun fetch() = service.spreadsheets().get(spreadsheetID).execute().sheets

    private val sheets by lazy { fetch() }

    private var sheet = sheets[0]

    private val requests = mutableListOf<Request>()

    var sheetName: String
        get() = sheet.properties.title
        set(value) {
            swapSheet(value)
        }

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

    private fun swapSheet(name: String) {
        sheets.forEach {
            if (it.properties.title == name) {
                sheet = it
                return
            }
        }
        throw IllegalArgumentException("Invalid sheet name: \"$name\"")
    }

    operator fun get(address: String): Array<Array<String>> {
        flush()
        val response: ValueRange = service.spreadsheets().values()
            .get(spreadsheetID, address)
            .execute()
        return response.getValues()?.map {
            it.map { that -> that.toString() }.toTypedArray()
        }?.toTypedArray() ?: emptyArray<Array<String>>()
    }

    operator fun String.get(address: String): Array<Array<String>> {
        sheetName = this
        return this@Workbook[address]
    }

    operator fun set(address: String, verticalOffset: Int = 0, horizontalOffset: Int = 0, value: String) =
        set(address, verticalOffset, horizontalOffset, arrayOf(arrayOf(value)))

    operator fun set(address: String, verticalOffset: Int = 0, horizontalOffset: Int = 0, values: Array<String>) {
        this[address, verticalOffset, horizontalOffset] =
            if (Range(address).height == 0) arrayOf(values)
            else Array(values.size) { arrayOf(values[it]) }
    }

    operator fun set(address: String, verticalOffset: Int = 0, horizontalOffset: Int = 0, values: Array<Array<String>>) {
        val targetRange = Range(address) + Cell(horizontalOffset, verticalOffset)
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

    operator fun set(address: String, verticalOffset: Int = 0, horizontalOffset: Int = 0, format: Format) {
        val targetRange = Range(address) + Cell(horizontalOffset, verticalOffset)
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

    operator fun set(cell: Cell, verticalOffset: Int = 0, horizontalOffset: Int = 0, value: String) = set(cell.toString(), verticalOffset, horizontalOffset, value)

    operator fun set(cell: Cell, verticalOffset: Int = 0, horizontalOffset: Int = 0, values: Array<String>) = set(cell.toString(), verticalOffset, horizontalOffset, values)

    operator fun set(address: Cell, verticalOffset: Int = 0, horizontalOffset: Int = 0, values: Array<Array<String>>) = set(address.toString(), verticalOffset, horizontalOffset, values)

    operator fun set(address: Cell, verticalOffset: Int = 0, horizontalOffset: Int = 0, format: Format) = set(address.toString(), verticalOffset, horizontalOffset, format)

    sealed interface Linearity
    data object ROWS : Linearity
    data object COLUMNS : Linearity

    operator fun Linearity.plusAssign(intRange: IntRange) {
        requests.add(Request().apply {
            insertDimension = InsertDimensionRequest().apply {
                range = DimensionRange().apply {
                    sheetId = this@Workbook.sheet.properties.sheetId
                    dimension = if (this@plusAssign == ROWS) "ROWS" else "COLUMNS"
                    startIndex = intRange.first
                    endIndex = intRange.last
                }
                inheritFromBefore = false
            }
        })
    }

    operator fun Linearity.minusAssign(intRange: IntRange) {
        requests.add(Request().apply {
            deleteDimension = DeleteDimensionRequest().setRange(
                DimensionRange().apply {
                    sheetId = this@Workbook.sheet.properties.sheetId
                    dimension = if (this@minusAssign == ROWS) "ROWS" else "COLUMNS"
                    startIndex = intRange.first
                    endIndex = intRange.last
                }
            )
        })
    }

    fun addCheckbox(address: String, verticalOffset: Int = 0, horizontalOffset: Int = 0) {
        requests.add(Request().apply {
            setDataValidation = SetDataValidationRequest().apply {
                range = Range(address).plus(Cell(horizontalOffset, verticalOffset)).gridRange(sheet.properties.sheetId)
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

    fun String.addCheckbox(address: String, verticalOffset: Int = 0, horizontalOffset: Int = 0) {
        sheetName = this
        this@Workbook.addCheckbox(address, verticalOffset, horizontalOffset)
    }

    fun mergeCells(address: String, verticalOffset: Int = 0, horizontalOffset: Int = 0) {
        requests.add(Request().apply {
            mergeCells = MergeCellsRequest().apply {
                range = Range(address).plus(Cell(horizontalOffset, verticalOffset)).gridRange(sheet.properties.sheetId)
                mergeType = "MERGE_ALL"
            }
        })
    }

    fun String.mergeCells(address: String, verticalOffset: Int = 0, horizontalOffset: Int = 0) {
        sheetName = this
        this@Workbook.mergeCells(address, verticalOffset, horizontalOffset)
    }

    fun Linearity.resize(intRange: IntRange, newWidth: Int? = null) {
        val dimensionRange = DimensionRange().apply {
            sheetId = sheet.properties.sheetId
            dimension = if (this@resize == ROWS) "ROWS" else "COLUMNS"
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

    fun protectRange(address: String, editors: Array<String>, verticalOffset: Int = 0, horizontalOffset: Int = 0) {
        requests.add(Request().apply {
            addProtectedRange = AddProtectedRangeRequest().apply {
                protectedRange = ProtectedRange().apply {
                    this.range = Range(address).plus(Cell(horizontalOffset, verticalOffset)).gridRange(sheet.properties.sheetId)
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
