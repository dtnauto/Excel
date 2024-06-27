package com.example.excel.usecase52

import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook

fun deleteRowsContainingSubstring(workbook: XSSFWorkbook, sheetName: String, column: Int, searchString: String) {
    val sheet = workbook.getSheet(sheetName)
    var lastRow = sheet.lastRowNum

    // Duyệt từ dưới lên trên để tránh vấn đề khi xóa hàng
    for (currentRow in lastRow downTo 0) {
        val cell = sheet.getRow(currentRow)?.getCell(column)

        // Kiểm tra nếu ô tồn tại và chứa chuỗi tìm kiếm
        if (cell != null && cell.toString().contains(searchString)) {
            sheet.removeRow(sheet.getRow(currentRow))
            if (currentRow != lastRow) {
                sheet.shiftRows(currentRow + 1, lastRow, -1)
            }
            lastRow-- // Giảm giá trị của lastRow vì một hàng đã bị xóa
        }
    }
}


fun newRowIf(workbook: XSSFWorkbook, sheetName: String, ranges: Array<Int>, insertItems: Array<String>, insertValues: Array<String>) {
    val sheet = workbook.getSheet(sheetName)
    var lastRow = sheet.lastRowNum

//    val newRow = 1
//
//    // Dịch chuyển các hàng bên dưới xuống
//    if (newRow <= sheet.lastRowNum) {
//        sheet.shiftRows(newRow, sheet.lastRowNum, 1)
//    }
//
//    // Chèn hàng mới
//    val row = sheet.createRow(newRow)
//
//    // Sao chép định dạng từ hàng phía trên (nếu có)
//    val sourceRow = sheet.getRow(newRow - 1)
//    if (sourceRow != null) {
//        copyRowStyles(sourceRow, row)
//    }

    var currentRow = sheet.firstRowNum
    while (currentRow <= lastRow) {
        val cellValue = sheet.getRow(currentRow)?.getCell(ranges[0])?.toString() ?: ""

        // Kiểm tra điều kiện để chèn dòng mới
        if (cellValue.contains("ahihi") ) {//&& cellValue.contains("] screen")) {
            var newRow = currentRow + 1

            // Chèn dòng mới với các giá trị từ mảng insertItems và insertValues
            for (item in insertItems) {
                for (value in insertValues) {
                    if (item.isNotBlank()) {
                        sheet.shiftRows(newRow, lastRow+1, 1)
                        val row = sheet.createRow(newRow)

                        // Thay thế giá trị chèn
                        val newCellValue = value.replace("[mediumItem]", "[$item]")
                        row.createCell(ranges[1]).setCellValue(newCellValue)
                        newRow++
                        lastRow++
                    }
                }
            }

            // Cập nhật lại currentRow để bỏ qua các dòng vừa chèn
            currentRow = newRow - 1
        }

        currentRow++
    }
}

// Hàm sao chép định dạng từ hàng nguồn sang hàng đích
fun copyRowStyles(sourceRow: Row, targetRow: Row) {
    for (cellNum in sourceRow.firstCellNum until sourceRow.lastCellNum) {
        val sourceCell = sourceRow.getCell(cellNum)
        if (sourceCell != null) {
            val newCell = targetRow.createCell(cellNum)
            newCell.cellStyle = sourceCell.cellStyle
        }
    }
}

fun cutAndFormatString(
    cellValue: String?,
    startKeyword: String = "",
    endKeyword: String = "",
    offsetLen: Int = 0,
    format: Int = 0,
    takeChars: Int = 0
): String? {
    if (cellValue.isNullOrEmpty()) {
        return null
    }

    var startPos: Int
    var endPos: Int
    var cutString: String
    var formattedString: String

    // Nếu startKeyword được truyền vào
    startPos = if (startKeyword.isNotEmpty()) {
        val offset = if (offsetLen == -1) startKeyword.length else offsetLen
        val startIndex = cellValue.indexOf(startKeyword, ignoreCase = true)
        if (startIndex != -1) startIndex + offset else 0
    } else {
        // Nếu không có startKeyword, bắt đầu từ đầu chuỗi
        0
    }

    // Nếu có endKeyword
    endPos = if (endKeyword.isNotEmpty()) {
        val endIndex = cellValue.indexOf(endKeyword, startPos, ignoreCase = true)
        if (endIndex != -1) endIndex else cellValue.length
    } else {
        // Nếu không có endKeyword, cắt đến cuối chuỗi
        cellValue.length
    }

    // Nếu endKeyword được tìm thấy
    cutString = if (endPos > startPos) {
        cellValue.substring(startPos, endPos)
    } else {
        // Nếu không tìm thấy endKeyword, trả về phần của chuỗi từ startPos đến cuối chuỗi
        cellValue.substring(startPos)
    }

    // Lấy x ký tự từ chuỗi được cắt ra nếu takeChars > 0
    if (takeChars > 0 && cutString.length > takeChars) {
        cutString = cutString.substring(0, takeChars)
    }

    formattedString = when (format) {
        1 -> {
            // Chuyển chuỗi thành chữ thường và viết hoa chữ cái đầu tiên
            cutString.lowercase().replace("_", " ").replaceFirstChar { it.uppercase() }.trim()
        }
        2 -> {
            // Chuyển chuỗi thành chữ thường, bỏ dấu _ và viết hoa chữ cái đầu tiên sau mỗi dấu _
            cutString.lowercase()
                .split("_")
                .joinToString("") { it.replaceFirstChar { char -> char.uppercase() } }
                .trim()
        }
        else -> {
            // Nếu không cần định dạng, trả về chuỗi cắt
            cutString
        }
    }

    // Trả về kết quả
    return formattedString
}
