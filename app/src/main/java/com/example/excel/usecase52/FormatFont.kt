package com.example.excel.usecase52

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFWorkbook

fun applyConditionalFormula(workbook: XSSFWorkbook, sheetName: Array<String>, ranges: Array<Int>) {
    val sheet = workbook.getSheet(sheetName[0])

    // Giả sử rằng người dùng đã chọn cột A và B trong VBA, chúng ta sẽ sử dụng cột A và B.

    val lastRow = sheet.lastRowNum
    var currentRow = 1

    while (currentRow <= lastRow) {
        val sourceCell = sheet.getRow(currentRow)?.getCell(ranges[0])

        // Kiểm tra nếu ô A là null hoặc trống
        if (sourceCell == null || sourceCell.cellType == org.apache.poi.ss.usermodel.CellType.BLANK) {
            currentRow++
            continue
        }

        val currentCellValue = sourceCell.toString()
        val previousCellValue = sheet.getRow(currentRow - 1)?.getCell(ranges[0])?.toString() ?: ""

        if (currentCellValue != previousCellValue) {
            if (currentCellValue.isEmpty()) {
                sourceCell.setCellValue(previousCellValue)  // nếu rỗng thì gán lại giá trị trước đó
            } else {
                val currentTargetCell = sheet.getRow(currentRow).getCell(ranges[1]) ?: sheet.getRow(currentRow).createCell(ranges[1])
                val previousTargetCell = sheet.getRow(currentRow - 1)?.getCell(ranges[1]) ?: sheet.getRow(currentRow).createCell(ranges[1])

                currentTargetCell.setCellValue(previousTargetCell.numericCellValue + 1)

                // Tạo đối tượng font
                val font = workbook.createFont().apply {
                    color = IndexedColors.BLACK.index // Màu chữ
                    // Các thuộc tính font khác có thể thiết lập ở đây, như đậm, nghiêng, cỡ chữ, vv.
                }

                // Áp dụng font vào cell style
                val cellStyle = workbook.createCellStyle().apply {
                    setFont(font)
                    setBorderTop(BorderStyle.THIN)
                }
                currentTargetCell.cellStyle = cellStyle
            }
        } else {
            val currentTargetCell = sheet.getRow(currentRow).getCell(ranges[1]) ?: sheet.getRow(currentRow).createCell(ranges[1])
            val previousTargetCell = sheet.getRow(currentRow - 1)?.getCell(ranges[1]) ?: sheet.getRow(currentRow).createCell(ranges[1])

            currentTargetCell.setCellValue(previousTargetCell.numericCellValue)

            // Tạo đối tượng font
            val font = workbook.createFont().apply {
                color = IndexedColors.GREY_25_PERCENT.index // Màu chữ
                // Các thuộc tính font khác có thể thiết lập ở đây
            }

            // Áp dụng font vào cell style
            val cellStyle = workbook.createCellStyle().apply {
                setFont(font)
                setBorderTop(BorderStyle.NONE)
            }
            currentTargetCell.cellStyle = cellStyle
        }

        currentRow++
    }
}