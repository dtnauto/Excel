package com.example.excel.usecase52

import com.example.excel.excelfile.columnNameToInt
import com.example.excel.excelfile.openExcelFile
import com.example.excel.excelfile.saveExcelFile
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.DataInput


fun main() {
    val filePath =
        "C:\\Users\\daotr\\Desktop\\New Microsoft Excel Worksheet.xlsx"
//        "D:\\working\\safety setting\\eng5.2\\New template.xlsx"

    val workbook = openExcelFile(filePath)
    if (workbook != null) {

        // thay doi mau chu
        /*applyConditionalFormula(workbook, arrayOf("Sheet1"), arrayOf(columnNameToInt("d"),
            columnNameToInt("d")
        ))
        applyConditionalFormula(workbook, arrayOf("Sheet1"), arrayOf(columnNameToInt("e"),
            columnNameToInt("e")
        ))
        applyConditionalFormula(workbook, arrayOf("Sheet1"), arrayOf(columnNameToInt("g"),
            columnNameToInt("f")
        ))
        applyConditionalFormula(workbook, arrayOf("Sheet1"), arrayOf(columnNameToInt("g"),
            columnNameToInt("g")
        ))*/

        // Gọi hàm để xóa các dòng
//        deleteRowsContainingSubstring(workbook, "Sheet4", 1, "bigItem")

        // Gọi hàm để chèn các dòng mới
//        addRowver2(workbook, "Sheet4")

        // Gọi hàm để chèn các dòng mới
        copyRowver2(workbook, "Sheet1")

//         xu ly overView
        for (i in 3..workbook.getSheet("Sheet1").lastRowNum) {
            val currentRow = workbook.getSheet("Sheet1").getRow(i)
            val inputCell = currentRow?.getCell(columnNameToInt("k"))?.stringCellValue
            if (inputCell != null) {
                overViewver2(inputCell)?.apply {
                    currentRow.createCell(columnNameToInt("n"))?.setCellValue(this)
                }
            }
        }

        // xu ly overView
//        for (i in 500..workbook.getSheet("Sheet1").lastRowNum) {
//            val currentRow = workbook.getSheet("Sheet1").getRow(i)
//            val inputCell = currentRow?.getCell(columnNameToInt("k"))?.stringCellValue
//            if (inputCell != null) {
//                overView(inputCell)?.apply {
//                    currentRow.createCell(columnNameToInt("n"))?.setCellValue(this)
//                }
//            }
//        }

        // ket thuc file
        saveExcelFile(workbook, filePath)
    }
}

fun overViewver2(inputCell: String?): String? {

    val fileContainUseCase =
//        "D:\\working\\safety setting\\eng5.2\\New template.xlsx"
        "C:\\Users\\daotr\\Desktop\\New Microsoft Excel Worksheet.xlsx"
    val workbookContainUseCase = openExcelFile(fileContainUseCase)!!
    val sheetContainUseCase = workbookContainUseCase.getSheet("UseCase")


    // Kiểm tra nếu inputCell là null hoặc chuỗi trống
    if (inputCell.isNullOrEmpty()) {
        return null
    }

    val extractedText = cutAndFormatString(
        inputCell,
        "[tempUseCase]",
        "[/tempUseCase]",
        -1
    )?.let {
        sheetContainUseCase.getRow(it.toInt())
            .getCell(columnNameToInt("i")).stringCellValue
            .replace(
                "[item][/item]",
                "<"
                        + (cutAndFormatString(
                    inputCell,
                    "[itemJP]",
                    "[/itemJP]",
                    -1
                ) ?: "")
                        + "> ["
                        + (cutAndFormatString(
                    inputCell,
                    "[itemEN]",
                    "[/itemEN]",
                    -1
                ) ?: "")
                        + "] ("
                        + (cutAndFormatString(
                    inputCell,
                    "[idBigItem]",
                    "[/idBigItem]",
                    -1
                ) ?: "")
                        + ")"
            )
            .replace(
                "[bigItem][/bigItem]",
                cutAndFormatString(
                    inputCell,
                    "[bigItem]",
                    "[/bigItem]",
                    -1
                ) ?: ""
            )
    }

    return extractedText
}

fun copyRowver2(workbook: XSSFWorkbook, sheetName: String) {

    val sheet = workbook.getSheet(sheetName)

    val fileContainUseCase =
//        "D:\\working\\safety setting\\eng5.2\\New template.xlsx"
        "C:\\Users\\daotr\\Desktop\\New Microsoft Excel Worksheet.xlsx"
    val workbookContainUseCase = openExcelFile(fileContainUseCase)!!
    val sheetContainUseCase = workbookContainUseCase.getSheet("Sheet4")

    var lastRowUseCase = sheetContainUseCase.lastRowNum
    var currentRowUseCase = 2 // fix current row
    while (currentRowUseCase <= lastRowUseCase) {
        val cellValue =
            sheetContainUseCase.getRow(currentRowUseCase)
                ?.getCell(columnNameToInt("k"))?.stringCellValue
                ?: ""
        if (cellValue.contains("screen1")) {
            break
        }
        currentRowUseCase++
    }

    var lastRow = sheet.lastRowNum
    var currentRow = 2 // fix current row
    while (currentRow <= lastRow) {
        val cellValue =
            sheet.getRow(currentRow)?.getCell(columnNameToInt("k"))?.stringCellValue ?: ""
        if (cellValue.contains("screen1")) {
            while (currentRowUseCase <= lastRowUseCase) {
                currentRow++
                currentRowUseCase++
                val cellUseCase =
                    sheetContainUseCase.getRow(currentRowUseCase)?.getCell(columnNameToInt("k"))
                if (cellUseCase != null) {
                    val sheetRow = sheet.getRow(currentRow) ?: sheet.createRow(currentRow)
                    sheetRow.createCell(columnNameToInt("k"))
                        .setCellValue(cellUseCase.stringCellValue)
                }
            }
            break
        }
        currentRow++
    }
}

fun addRowver2(workbook: XSSFWorkbook, sheetName: String) {
    val sheet = workbook.getSheet(sheetName)

    val fileContainItem =
//        "D:\\working\\safety setting\\eng5.2\\New template.xlsx"
        "C:\\Users\\daotr\\Desktop\\New Microsoft Excel Worksheet.xlsx"
    val workbookContainItem = openExcelFile(fileContainItem)!!
    val sheetContainItem = workbookContainItem.getSheet("Item")

    val fileContainUseCase =
//        "D:\\working\\safety setting\\eng5.2\\New template.xlsx"
        "C:\\Users\\daotr\\Desktop\\New Microsoft Excel Worksheet.xlsx"
    val workbookContainUseCase = openExcelFile(fileContainUseCase)!!
    val sheetContainUseCase = workbookContainUseCase.getSheet("UseCase")

    fun getRanges(
        // get range từng khoảng giống nhau
        sheetGetRanges: XSSFSheet,
        startRow: Int,
        endRow: Int,
        columnGetRanges: Int,
    ): MutableList<IntRange> {
        val ranges = mutableListOf<IntRange>()

        var start = startRow
        var currentCellValue: String? = null

        for (row in startRow..endRow) {
            val cellValue =
                sheetGetRanges.getRow(row).getCell(columnGetRanges).stringCellValue
            if (currentCellValue == null) {
                currentCellValue = cellValue
                start = row
            } else if (cellValue != currentCellValue) {
                ranges.add(start until row)
                currentCellValue = cellValue
                start = row
            }
        }
        ranges.add(start..endRow)  // Add the last range

        return ranges
    }


    var lastRow = sheet.lastRowNum
    var currentRow = 2 // fix current row
    fun createNewCellValue(
        rowUseCase: Int,
        rowItem: Int,
        extraText: String = ""
    ): String {
        var cellValueTemplate =
            "[useCase]" +
                    sheetContainUseCase.getRow(rowUseCase)
                        .getCell(columnNameToInt("h")).stringCellValue +
                    "[/useCase]" +
                    "[tempUseCase]" +
                    rowUseCase +
                    "[/tempUseCase]" +
                    "[bigItem]" + cutAndFormatString(
                sheetContainItem.getRow(rowItem).getCell(columnNameToInt("c")).stringCellValue,
                "[",
                "]",
                1
            ) + "[/bigItem]" +
                    "[itemJP]" + cutAndFormatString(
                sheetContainItem.getRow(rowItem).getCell(columnNameToInt("d")).stringCellValue,
                "[",
                "]",
                1
            ) + "[/itemJP]" +
                    "[itemEN]" + cutAndFormatString(
                sheetContainItem.getRow(rowItem).getCell(columnNameToInt("e")).stringCellValue,
                "[",
                "]",
                1
            ) + "[/itemEN]" +
                    "[idBigItem]" + cutAndFormatString(
                sheetContainItem.getRow(rowItem).getCell(columnNameToInt("f")).stringCellValue,
                "[",
                "]",
                1
            ) + "[/idBigItem]"

        return cellValueTemplate
    }

    fun shiftAndCreateRow(newCellValue: String) {
        sheet.shiftRows(++currentRow, ++lastRow, 1)
        sheet.createRow(currentRow).createCell(columnNameToInt("k")).setCellValue(newCellValue)
    }

    fun processRows(
        rowItemRange: IntRange,
        rowUseCaseRange: IntRange,
        ifAction: (Int) -> Boolean = { _ -> false },
        extraText: (Int, Int) -> String = { _, _ -> "" }
    ) {
        for (rowItem in rowItemRange) {
            if (ifAction(rowItem)) {
                for (rowUseCase in rowUseCaseRange) {
                    shiftAndCreateRow( // ki tự
                        createNewCellValue(
                            rowUseCase,
                            rowItem,
                            extraText(rowItem, rowUseCase)
                        ) // chen cai gi
                    )
                }
            }
        }
    }
    while (currentRow <= lastRow) {
        val cellValue =
            sheet.getRow(currentRow)?.getCell(columnNameToInt("k"))?.stringCellValue ?: ""
        if (cellValue.contains("screen1")) {
            val rangesOfItem = getRanges(sheetContainItem, 2, 20, columnNameToInt("c"))
            println(rangesOfItem)
            val rangesOfUseCase = getRanges(sheetContainUseCase, 2, 35, columnNameToInt("g"))
            println(rangesOfUseCase)
            for (rangeOfItem in rangesOfItem) {
                //display
                processRows(rangeOfItem, rangesOfUseCase[0],
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] != 'A' || cellAction.length < 7)
                    },
                    extraText = { _, _ ->
                        ""
                    })

                processRows(rangeOfItem, rangesOfUseCase[0].first + 1..rangesOfUseCase[0].last,
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 1)
                    },
                    extraText = { _, _ ->
                        ""
                    })

                //update
                processRows(rangeOfItem, rangesOfUseCase[1],
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] != 'A' || cellAction.length < 7)
                    },
                    extraText = { _, _ ->
                        ""
                    })

                processRows(rangeOfItem, rangesOfUseCase[1],
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 1)
                    },
                    extraText = { _, _ ->
                        ""
                    })

                // change
                processRows(rangeOfItem, rangesOfUseCase[2],
                    ifAction = { rowItem ->
                        val cellAction = sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == '0')
                    },
                    extraText = { _, _ ->
                        ""
                    })

                processRows(rangeOfItem, rangesOfUseCase[3],
                    ifAction = { rowItem ->
                        val cellAction = sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.isEmpty())

                    },
                    extraText = { _, _ ->
                        ""
                    })

                //reset
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(0)..rangesOfUseCase[4].elementAt(0),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 1)
                    },
                    extraText = { _, _ ->
                        ""
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(2)..rangesOfUseCase[4].elementAt(2),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 2)
                    },
                    extraText = { _, _ ->
                        ""
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(3)..rangesOfUseCase[4].elementAt(3),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 1)
                    },
                    extraText = { _, _ ->
                        ""
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(1)..rangesOfUseCase[4].elementAt(1),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 1)
                    },
                    extraText = { _, _ ->
                        ""
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(4)..rangesOfUseCase[4].elementAt(4),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 3)
                    },
                    extraText = { _, _ ->
                        ""
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(5)..rangesOfUseCase[4].elementAt(5),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 0)
                    },
                    extraText = { _, _ ->
                        ""
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(6)..rangesOfUseCase[4].elementAt(6),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 3)
                    },
                    extraText = { _, _ ->
                        ""
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(7)..rangesOfUseCase[4].elementAt(7),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 0)
                    },
                    extraText = { _, _ ->
                        ""
                    })

                //back
                processRows(rangeOfItem, rangesOfUseCase[5],
                    ifAction = { rowItem ->
                        rowItem == rangeOfItem.last()
                    },
                    extraText = { _, _ ->
                        ""
                    })
                //////////////////////////////
            }
            break
        }
        currentRow++
    }
}