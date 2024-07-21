package com.example.excel.usecase52

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook

fun applyConditionalFormula(workbook: XSSFWorkbook, sheetName: String, columnRef: Int, columnChange: Int, currentRow: Int = 0, lastRow: Int = 0) {
    val sheet = workbook.getSheet(sheetName)

    val lastRow = if (lastRow == 0) sheet.lastRowNum else lastRow
    var currentRow = if (currentRow == 0) sheet.firstRowNum else currentRow  //fix lại chỉ số đầu

    while (currentRow <= lastRow) {

        val currentCellValue = sheet.getRow(currentRow)?.getCell(columnRef)?.toString() ?: ""
        val previousCellValue = sheet.getRow(currentRow - 1)?.getCell(columnRef)?.toString() ?: ""

        if (currentCellValue != previousCellValue) {
            if (currentCellValue.isEmpty()) {
                sheet.getRow(currentRow)?.getCell(columnRef)?.setCellValue(previousCellValue)  // nếu rỗng thì gán lại giá trị trước đó
            } else {
                val currentTargetCell = sheet.getRow(currentRow).getCell(columnChange) ?: sheet.getRow(currentRow).createCell(columnChange)

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
            val currentTargetCell = sheet.getRow(currentRow).getCell(columnChange) ?: sheet.getRow(currentRow).createCell(columnChange)

            // Tạo đối tượng font
            val font = workbook.createFont().apply {
//                color = IndexedColors.GREY_50_PERCENT.index // Màu chữ
                // Các thuộc tính font khác có thể thiết lập ở đây
                val rgb = XSSFColor(byteArrayOf(174.toByte(),170.toByte(),170.toByte()))
                (this as XSSFFont).setColor(rgb)
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

