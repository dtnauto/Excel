package com.example.excel

import com.example.excel.excelfile.openExcelFile
import com.example.excel.excelfile.saveExcelFile
import com.example.excel.screenhmi52.processColumnScreenSpec
import com.example.excel.usecase52.processColumnSheet4
import org.apache.poi.xssf.usermodel.XSSFWorkbook


fun updateExcelSheetWithSumFormula(workbook: XSSFWorkbook) {
    val sheet = workbook.getSheetAt(0) // Mở sheet đầu tiên

    // Chèn công thức SUM vào cột C từ hàng 5 đến hàng 10
    for (rowNum in 4..9) {  // Hàng 5 đến hàng 10 trong Excel (index 4 đến 9)
        val row = sheet.getRow(rowNum) ?: sheet.createRow(rowNum)
        val cellA = row.getCell(0) ?: row.createCell(0).apply { setCellValue(0.0) }
        val cellB = row.getCell(1) ?: row.createCell(1).apply { setCellValue(0.0) }
        val cellC = row.createCell(2)
        cellC.cellFormula = "SUM(A${rowNum + 1}, B${rowNum + 1})"
    }
}


fun updateExcelFileWithSumFormula(filePath: String) {
    val workbook = openExcelFile(filePath)
    if (workbook != null) {
        updateExcelSheetWithSumFormula(workbook)
        saveExcelFile(workbook, filePath)
    }
}

fun main() {
    val filePath = "C:\\Users\\daotr\\Desktop\\New Microsoft Excel Worksheet.xlsx"
//    updateExcelFileWithSumFormula(filePath)
//    applyConditionalFormula(filePath)
    val workbook = openExcelFile(filePath)
    if (workbook != null) {
//        applyConditionalFormula(workbook, arrayOf("Sheet1"), arrayOf(0,1))
//        deleteRowsContainingSubstring(workbook, "Sheet4", 1, "mediumItem")

        /////////
        /*val sheetName = "Sheet4"
        val ranges = arrayOf(0, 1, 2) // Chỉ mục của các cột (zero-based index)
        val insertItems = arrayOf("OFF", "WEAK", "MEDIUM", "STRONG") // Mảng giá trị cần chèn
        val insertValues = arrayOf(
            "Display [mediumItem] setting item",
            "Update the display of [mediumItem] setting item",
            "Change setting of [mediumItem] by user operation"
        ) // Mảng giá trị thay thế

        // Gọi hàm để chèn các dòng mới
        newRowIf(workbook, sheetName, ranges, insertItems, insertValues)*/

        /*processColumnSheet4(workbook, "Sheet5", 5)
        processColumnSheet4(workbook, "Sheet5", 6)
        processColumnSheet4(workbook, "Sheet5", 7)*/

        processColumnScreenSpec(workbook, "Sheet6", 0)

        saveExcelFile(workbook, filePath)
    }
}

