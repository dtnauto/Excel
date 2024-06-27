package com.example.excel.screenhmi52

import com.example.excel.excelfile.columnNameToInt
import com.example.excel.excelfile.openExcelFile
import com.example.excel.excelfile.saveExcelFile
import com.example.excel.usecase52.cutAndFormatString
import com.example.excel.usecase52.overView
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream

fun main() {
    val filePathSource =
//        "D:\\working\\safety setting\\Input\\Spec\\ScreenSpec\\画面仕様参照ツール_ver2.01.xlsm"
        "D:\\working\\safety setting\\Input\\Spec\\ScreenTransitionSpec\\CDC-HST-004-003-03_画面遷移仕様書(MCStd-Setting-安全装備設定) (3).xlsm"

    val filePathDestination =
        "D:\\working\\safety setting\\Input\\Spec\\HMI_Screen_NHANDT53 - Copy.xlsx"
    val sourceWorkbook = openExcelFile(filePathSource)
    val destinationWorkbook = openExcelFile(filePathDestination)
    if (sourceWorkbook != null && destinationWorkbook != null) {

//        createScreenSpec(sourceWorkbook, destinationWorkbook, "Sheet1", columnNameToInt("b"))

        processColumnScreenSpec(arrayOf(destinationWorkbook),"Sheet1", columnNameToInt("d"))
//        linkScreenTransitionSpec(arrayOf(destinationWorkbook,sourceWorkbook), arrayOf("Sheet1"))
        saveExcelFile(sourceWorkbook, filePathSource)
        saveExcelFile(destinationWorkbook, filePathDestination)
    }
}

fun processColumnScreenSpec(workbook: Array<XSSFWorkbook>, sheetName: String, columnToProcess: Int, firstRow: Int = -1) {
    val sheet = workbook[0].getSheet(sheetName)

    val firstRow = if (firstRow == -1) sheet.firstRowNum else 0
    val lastRow = sheet.lastRowNum

    val rows = (firstRow..lastRow).toList().toIntArray()
    val columnSource = columnNameToInt("c")

    for (i in rows) {
        val currentRow = sheet.getRow(i)
        val inputCell = currentRow?.getCell(columnSource)?.stringCellValue
        if (inputCell != null) {
            when (columnToProcess) {
                columnNameToInt("a") -> {
                    cutAndFormatString(inputCell, takeChars = 7)?.apply {
                        currentRow.createCell(columnToProcess)?.setCellValue(this)
                    }
                }
                columnNameToInt("c"), columnNameToInt("f") -> {
                    cutAndFormatString(inputCell, takeChars = 7)?.apply {
                        currentRow.createCell(columnToProcess)?.setCellValue(this)
                    }
                }
                columnNameToInt("d") -> {
                    cutAndFormatString(inputCell,"[","]",1)?.apply {
                        currentRow.createCell(columnToProcess)?.setCellValue(this)
                    }
                }
            }

        }
    }
}


fun createScreenSpec(
    sourceWorkbook: XSSFWorkbook,
    destinationWorkbook: XSSFWorkbook,
    destinationSheetName: String,
    columnToProcess: Int,
) {
    // Lấy sheet của workbook destination
    val destinationSheet = destinationWorkbook.getSheet(destinationSheetName)

    // Bắt đầu ghi dữ liệu từ hàng 0 của sheet đích
    var destinationRowIndex = 4

    // Duyệt qua các sheet từ vị trí i đến hết của workbook 1
    for (i in 3 until sourceWorkbook.numberOfSheets) {
        val sourceSheet = sourceWorkbook.getSheetAt(i)

        // Duyệt qua các hàng từ hàng 12 đến hàng cuối cùng
        for (rowIndex in 11..sourceSheet.lastRowNum) {
            val sourceRow = sourceSheet.getRow(rowIndex)
            val sourceCell = sourceRow?.getCell(columnNameToInt("f"))?.stringCellValue
            // Sao chép giá trị từ ô nguồn sang ô đích
            if (sourceCell != null) {
                val destinationRow = destinationSheet.getRow(destinationRowIndex++)
                val destinationCell = destinationRow.createCell(columnToProcess)
                destinationCell.setCellValue(sourceCell.toString())
            }
        }
    }
}

fun linkScreenTransitionSpec(
    workbook: Array<XSSFWorkbook>,
    sheetName: Array<String>,
) {
    // Lấy sheet của workbook destination
    val destinationSheet = workbook[0].getSheet(sheetName[0])

    // Bắt đầu ghi dữ liệu từ hàng 0 của sheet đích
    var destinationRowIndex = 4

    // Duyệt qua các sheet từ vị trí i đến hết của workbook 1
    for (i in destinationRowIndex .. destinationSheet.lastRowNum) {
        val destinationRow = destinationSheet.getRow(i)
        val destinationCellSheetName = destinationRow?.getCell(columnNameToInt("a"))?.stringCellValue
        val destinationCellPartID = destinationRow?.getCell(columnNameToInt("b"))?.stringCellValue
        if (destinationCellSheetName != null && destinationCellPartID != null) {
            val sourceSheet = workbook[1].getSheet(destinationCellSheetName)
            if (sourceSheet!= null){
                for (j in 3 .. sourceSheet.lastRowNum) {
                    val sourceRow = sourceSheet.getRow(j)
                    val sourceRowAction = sourceSheet.getRow(j + 1)
                    val sourceCellPartID = sourceRow?.getCell(columnNameToInt("o"))?.stringCellValue
                    val sourceCellName = sourceRow?.getCell(columnNameToInt("q"))?.stringCellValue
                    val sourceCellAction = sourceRowAction?.getCell(columnNameToInt("ab"))?.stringCellValue
                    // Sao chép giá trị từ ô nguồn sang ô đích
                    if (sourceCellPartID != null && sourceCellName != null && sourceCellAction != null && sourceCellPartID.contains(destinationCellPartID)) {
                        destinationRow.createCell(columnNameToInt("c")).setCellValue(sourceCellName)
                        destinationRow.createCell(columnNameToInt("f")).setCellValue(sourceCellAction)
                    }
                }
            }
        }
    }
}