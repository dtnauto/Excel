package com.example.excel.screenhmi52

import com.example.excel.excelfile.columnNameToInt
import com.example.excel.usecase52.cutAndFormatString
import com.example.excel.usecase52.overView
import org.apache.poi.xssf.usermodel.XSSFWorkbook

fun processColumnScreenSpec(workbook: XSSFWorkbook, sheetName: String, columnToProcess: Int) {
    val sheet = workbook.getSheet(sheetName)

    val firstRow = sheet.firstRowNum
    val lastRow = sheet.lastRowNum

    val rows = (firstRow..lastRow).toList().toIntArray()
    val columnSource = columnNameToInt("b")

    for (i in rows) {
        val currentRow = sheet.getRow(i)
        val inputCell = currentRow?.getCell(columnSource)?.stringCellValue
        if (inputCell != null) {
            when (columnToProcess) {
                0 ->{
                    cutAndFormatString(inputCell)?.apply {
                        currentRow.createCell(columnToProcess)?.setCellValue(this)
                    }
                }
            }

        }
    }
}