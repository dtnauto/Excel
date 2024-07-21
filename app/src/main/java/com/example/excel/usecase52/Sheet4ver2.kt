package com.example.excel.usecase52

import com.example.excel.excelfile.columnNameToInt
import com.example.excel.excelfile.openExcelFile
import com.example.excel.excelfile.saveExcelFile
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.DataInput

const val file_Path =
    "C:\\Users\\daotr\\Desktop\\New Microsoft Excel Worksheet.xlsx"
//    "D:\\working\\safety setting\\eng5.2\\New template.xlsx"

const val file_ContainUseCase =
    "C:\\Users\\daotr\\Documents\\AndroidStudioProjects\\Excel"
//    "C:\\Users\\daotr\\Desktop\\HMI_Screen_NHANDT53 - Copy.xlsx"
////        "D:\\working\\safety setting\\Input\\Spec\\HMI_Screen_NHANDT53 - Copy.xlsx"

const val file_ContainItem =
    "C:\\Users\\daotr\\Documents\\AndroidStudioProjects\\Excel"
//"C:\\Users\\daotr\\Desktop\\HMI_Screen_NHANDT53 - Copy.xlsx"
//        "D:\\working\\safety setting\\Input\\Spec\\HMI_Screen_NHANDT53 - Copy.xlsx"

const val mark_text = "screen1screen2"

fun main() {
    val filePath = file_Path

    val workbook = openExcelFile(filePath)
    if (workbook != null) {
        // Gọi hàm để chèn các dòng mới
        addRowver2(workbook, "Sheet4")
        // Gọi hàm để xóa các dòng
//        deleteRowsContainingSubstring(workbook, "Sheet4", 1, "bigItem")

        // Gọi hàm để copy các usecase các dòng mới
        copyRowver2(workbook, "Sheet1")

        // replaceUseCase
        replaceUseCase(workbook, "Sheet1")

        // ket thuc file
        saveExcelFile(workbook, filePath)
    }
}

fun replaceUseCase(workbook: XSSFWorkbook, sheetName: String){
    val sheet = workbook.getSheet(sheetName)
    var lastRow = sheet.lastRowNum
    var currentRow = 2 // fix current row
    while (currentRow <= lastRow) {
        val cellValue =
            sheet.getRow(currentRow)?.getCell(columnNameToInt("k"))?.stringCellValue ?: ""
        if (cellValue.contains(mark_text)) {
            currentRow++
            for (row in currentRow..sheet.lastRowNum) {
                val inputCell =
                    sheet.getRow(row)?.getCell(columnNameToInt("k"))?.stringCellValue
                if (inputCell != null) {
//                    replaceText(inputCell, columnNameToInt("e"))?.apply {
//                        sheet.getRow(row).createCell(columnNameToInt("e"))?.setCellValue(this)
//                    }
//
//                    sheet.getRow(row).createCell(columnNameToInt("f"))?.cellFormula =
//                        "IF(E${row + 1}<>E${row},1,I" +
//                                "F(G${row + 1}<>G${row},F${row}+1,F${row}))"
//
//                    replaceText(inputCell, columnNameToInt("g"))?.apply {
//                        sheet.getRow(row).createCell(columnNameToInt("g"))?.setCellValue(this)
//                    }
//
//                    replaceText(inputCell, columnNameToInt("h"))?.apply {
//                        sheet.getRow(row).createCell(columnNameToInt("h"))?.setCellValue(this)
//                    }
//
//                    replaceText(inputCell, columnNameToInt("i"))?.apply {
//                        sheet.getRow(row).createCell(columnNameToInt("i"))?.setCellValue(this)
//                    }
//
//                    sheet.getRow(row).createCell(columnNameToInt("j"))?.cellFormula =
//                        "IF(F${row + 1}<>F${row},1,J${row}+1)"
//
//                    sheet.getRow(row).createCell(columnNameToInt("l"))?.cellFormula =
//                        "\"UC.\"&D${row + 1}&\"-\"&F${row + 1}&\"-\"&J${row + 1}"
//
//                    replaceText(inputCell, columnNameToInt("m"))?.apply {
//                        sheet.getRow(row).createCell(columnNameToInt("m"))?.setCellValue(this)
//                    }
//
//                    replaceText(inputCell, columnNameToInt("n"))?.apply {
//                        sheet.getRow(row).createCell(columnNameToInt("n"))?.setCellValue(this)
//                    }

                    // chu y fix UseCase de cuoi cung do
                    replaceText(inputCell, columnNameToInt("k"))?.apply {
                        sheet.getRow(row).createCell(columnNameToInt("k"))?.setCellValue(this)
                    }
                }
            }

            // thay doi mau chu
//            applyConditionalFormula(
//                workbook, sheetName,
//                columnNameToInt("d"), columnNameToInt("d"),
//                currentRow
//            )
//            applyConditionalFormula(
//                workbook, sheetName,
//                columnNameToInt("e"), columnNameToInt("e"),
//                currentRow
//            )
//            applyConditionalFormula(
//                workbook, sheetName,
//                columnNameToInt("g"), columnNameToInt("f"),
//                currentRow
//            )
//            applyConditionalFormula(
//                workbook, sheetName,
//                columnNameToInt("g"), columnNameToInt("g"),
//                currentRow
//            )

            break
        }
        currentRow++
    }
}

fun replaceText(inputCell: String?, columnRef: Int): String? {

    val fileContainUseCase = file_ContainUseCase
    val workbookContainUseCase = openExcelFile(fileContainUseCase)!!
    val sheetContainUseCase = workbookContainUseCase.getSheet("UseCase")


    // Kiểm tra nếu inputCell là null hoặc chuỗi trống
    if (inputCell.isNullOrEmpty()) {
        return null
    }

    val extractedText = cutAndFormatString(
        inputCell,
        "[rowUseCase]",
        "[/rowUseCase]",
        -1
    )?.toInt()?.let {
        sheetContainUseCase.getRow(it).getCell(columnRef).stringCellValue
            .replace(
                "[idBackItem][/idBackItem]",
                cutAndFormatString(
                    inputCell,
                    "[idBackItem]",
                    "[/idBackItem]",
                    -1
                )?.let { text -> if (text == "null") "" else text } ?: ""
            )
            .replace(
                "[backItem][/backItem]",
                cutAndFormatString(
                    inputCell,
                    "[backItem]",
                    "[/backItem]",
                    -1
                )?.let { text -> if (text == "null") "" else text } ?: ""
            )
            .replace(
                "[idBigItem][/idBigItem]",
                cutAndFormatString(
                    inputCell,
                    "[idBigItem]",
                    "[/idBigItem]",
                    -1
                )?.let { text -> if (text == "null") "" else text }
                    ?: ""
            )
            .replace(
                "[bigItem][/bigItem]",
                cutAndFormatString(
                    inputCell,
                    "[bigItem]",
                    "[/bigItem]",
                    -1
                )?.let { text -> if (text == "null") "" else text } ?: ""
            )
            .replace(
                "[itemJP][/itemJP]",
                cutAndFormatString(
                    inputCell,
                    "[itemJP]",
                    "[/itemJP]",
                    -1
                )?.let { text -> if (text == "null") "" else text } ?: ""
            )
            .replace(
                "[itemEN][/itemEN]",
                cutAndFormatString(
                    inputCell,
                    "[itemEN]",
                    "[/itemEN]",
                    -1
                )?.let { text -> if (text == "null") "" else text } ?: ""
            )
            .replace(
                "[itemIdScreen][/itemIdScreen]",
                cutAndFormatString(
                    inputCell,
                    "[itemIdScreen]",
                    "[/itemIdScreen]",
                    -1
                )?.let { text -> if (text == "null" || text == "NoChangeScreen" || text == "SmallScreen") "" else text }
                    ?: ""
            )
            //special ///////////
            .replace(
                "[itemJPButton][/itemJPButton]",
                (cutAndFormatString(cutAndFormatString(
                    inputCell,
                    "[itemJP]",
                    "[/itemJP]",
                    -1
                )?.let { text -> if (text == "null") "" else text } ?: "",
                    endKeyword = "リセット") ?: "").let { text -> if (text == "") "" else text + "リセット" }
            )
            .replace(
                "[itemENButton][/itemENButton]",
                (cutAndFormatString(cutAndFormatString(
                    inputCell,
                    "[itemEN]",
                    "[/itemEN]",
                    -1
                )?.let { text -> if (text == "null") "" else text } ?: "",
                    endKeyword = "reset") ?: "").let { text -> if (text == "") "" else text + "reset" }
            )
            .replace(
                "[bigItemSetting][/bigItemSetting]",
                (cutAndFormatString(
                    inputCell,
                    "[bigItem]",
                    "[/bigItem]",
                    -1
                )?.let { text -> if (text == "null") "" else text } ?: "")
                    .replace("画面", "")
                    .replace(" screen", "")
            )
            .replace("()", "")
            .replace("  ", " ")
    }

    return extractedText
}

fun copyRowver2(workbook: XSSFWorkbook, sheetName: String) {

    val sheet = workbook.getSheet(sheetName)

    val fileContainUseCase = file_Path
    val workbookContainUseCase = openExcelFile(fileContainUseCase)!!
    val sheetContainUseCase = workbookContainUseCase.getSheet("Sheet4")

    var lastRowUseCase = sheetContainUseCase.lastRowNum
    var currentRowUseCase = 2 // fix current row
    while (currentRowUseCase <= lastRowUseCase) {
        val cellValue =
            sheetContainUseCase.getRow(currentRowUseCase)
                ?.getCell(columnNameToInt("k"))?.stringCellValue
                ?: ""
        if (cellValue.contains(mark_text)) {
            break
        }
        currentRowUseCase++
    }

    var lastRow = sheet.lastRowNum
    var currentRow = 2 // fix current row
    while (currentRow <= lastRow) {
        val cellValue =
            sheet.getRow(currentRow)?.getCell(columnNameToInt("k"))?.stringCellValue ?: ""
        if (cellValue.contains(mark_text)) {
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

    val fileContainItem = file_ContainItem
    val workbookContainItem = openExcelFile(fileContainItem)!!
    val sheetContainItem = workbookContainItem.getSheet("Item")

    val fileContainUseCase = file_ContainUseCase
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
        var cellValueTemplate = "[rowUseCase]" +
                rowUseCase +
                "[/rowUseCase]" +
                "[idBackItem]" + cutAndFormatString(
            sheetContainItem.getRow(rowItem).getCell(columnNameToInt("b")).stringCellValue
        ) + "[/idBackItem]" +
                "[backItem]" + cutAndFormatString(
            sheetContainItem.getRow(rowItem).getCell(columnNameToInt("c")).stringCellValue
        ) + "[/backItem]" +
                "[idBigItem]" + cutAndFormatString(
            sheetContainItem.getRow(rowItem).getCell(columnNameToInt("d")).stringCellValue
        ) + "[/idBigItem]" +
                "[bigItem]" + cutAndFormatString(
            sheetContainItem.getRow(rowItem).getCell(columnNameToInt("e")).stringCellValue
        ) + "[/bigItem]" +
                "[itemJP]" + cutAndFormatString(
            sheetContainItem.getRow(rowItem).getCell(columnNameToInt("f")).stringCellValue,
            "[",
            "]",
            1
        ) + "[/itemJP]" +
                "[itemEN]" + cutAndFormatString(
            sheetContainItem.getRow(rowItem).getCell(columnNameToInt("g")).stringCellValue,
            "[",
            "]",
            1
        ) + "[/itemEN]" +
                "[itemIdScreen]" + cutAndFormatString(
            sheetContainItem.getRow(rowItem).getCell(columnNameToInt("h")).stringCellValue,
            "[",
            "]",
            1
        ) + "[/itemIdScreen]"

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
        if (cellValue.contains(mark_text)) {
            val rangesOfItem = getRanges(sheetContainItem, 53, 78, columnNameToInt("d"))
            println(rangesOfItem)
            val rangesOfUseCase = getRanges(sheetContainUseCase, 2, 35, columnNameToInt("g"))
            println(rangesOfUseCase)
            for (rangeOfItem in rangesOfItem) {
                //display item
                processRows(rangeOfItem, rangesOfUseCase[0],
                    ifAction = { rowItem ->
                        val cellidBigItem = sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("h")).stringCellValue

                        (cutAndFormatString(cellidBigItem, takeChars = 7) == "ST_SF_0"
                                && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() >= 10)
                                || cellidBigItem == "NoChangeScreen"
                    },
                    extraText = { _, _ ->
                        ""
                    })

                processRows(rangeOfItem, rangesOfUseCase[0].first + 1..rangesOfUseCase[0].last,
                    ifAction = { rowItem ->
                        val cellidBigItem = sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("h")).stringCellValue

                        (cutAndFormatString(cellidBigItem, takeChars = 7) == "ST_SF_0"
                                && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() > 1
                                && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() < 10)
                                || (cutAndFormatString(cellidBigItem, takeChars = 7) == "ST_SF_A"
                                && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() % 4 == 1)
                    },
                    extraText = { _, _ ->
                        ""
                    })

                //update item
                processRows(rangeOfItem, rangesOfUseCase[1],
                    ifAction = { rowItem ->
                        val cellidBigItem = sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("h")).stringCellValue

                        (cutAndFormatString(cellidBigItem, takeChars = 7) == "ST_SF_0"
                                && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() >= 10)
                                || cellidBigItem == "NoChangeScreen"
                    },
                    extraText = { _, _ ->
                        ""
                    })

                processRows(rangeOfItem, rangesOfUseCase[1].first + 1..rangesOfUseCase[1].last,
                    ifAction = { rowItem ->
                        val cellidBigItem = sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("h")).stringCellValue

                        (cutAndFormatString(cellidBigItem, takeChars = 7) == "ST_SF_0"
                                && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() > 1
                                && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() < 10)
                                || (cutAndFormatString(cellidBigItem, takeChars = 7) == "ST_SF_A"
                                && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() % 4 == 1)
                    },
                    extraText = { _, _ ->
                        ""
                    })

                // display value item
                processRows(rangeOfItem, rangesOfUseCase[2],
                    ifAction = { rowItem ->
                        val cellidBigItem = sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("h")).stringCellValue

                        cellidBigItem == "SmallScreen"
                    },
                    extraText = { _, _ ->
                        ""
                    })

                // update value item
                processRows(rangeOfItem, rangesOfUseCase[3],
                    ifAction = { rowItem ->
                        val cellidBigItem = sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("h")).stringCellValue

                        cellidBigItem == "SmallScreen"
                    },
                    extraText = { _, _ ->
                        ""
                    })

                // change
//                processRows(rangeOfItem, rangesOfUseCase[2],
//                    ifAction = { rowItem ->
//                        val cellidBigItem = sheetContainItem.getRow(rowItem)
//                            .getCell(columnNameToInt("h")).stringCellValue
//                        (cellidBigItem.length > 7 && cellidBigItem[6] == '0')
//                    },
//                    extraText = { _, _ ->
//                        ""
//                    })
//
//                processRows(rangeOfItem, rangesOfUseCase[3],
//                    ifAction = { rowItem ->
//                        val cellidBigItem = sheetContainItem.getRow(rowItem)
//                            .getCell(columnNameToInt("h")).stringCellValue
//                        (cellidBigItem.isEmpty())
//
//                    },
//                    extraText = { _, _ ->
//                        ""
//                    })

                //reset
                processRows(rangeOfItem,
                    rangesOfUseCase[6].elementAt(0)..rangesOfUseCase[6].elementAt(0),
                    ifAction = { rowItem ->
                        val cellidBigItem =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("h")).stringCellValue

                        (cutAndFormatString(cellidBigItem, takeChars = 7) == "ST_SF_A"
                                && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() % 4 == 1)
                    },
                    extraText = { _, _ ->
                        ""
                    })
//                processRows(rangeOfItem,
//                    rangesOfUseCase[4].elementAt(2)..rangesOfUseCase[4].elementAt(2),
//                    ifAction = { rowItem ->
//                        val cellidBigItem =
//                            sheetContainItem.getRow(rowItem)
//                                .getCell(columnNameToInt("h")).stringCellValue
//                        (cellidBigItem.length > 7 && cellidBigItem[6] == 'A' && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() % 4 == 2)
//                    },
//                    extraText = { _, _ ->
//                        ""
//                    })
//                processRows(rangeOfItem,
//                    rangesOfUseCase[4].elementAt(3)..rangesOfUseCase[4].elementAt(3),
//                    ifAction = { rowItem ->
//                        val cellidBigItem =
//                            sheetContainItem.getRow(rowItem)
//                                .getCell(columnNameToInt("h")).stringCellValue
//                        (cellidBigItem.length > 7 && cellidBigItem[6] == 'A' && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() % 4 == 1)
//                    },
//                    extraText = { _, _ ->
//                        ""
//                    })
//                processRows(rangeOfItem,
//                    rangesOfUseCase[4].elementAt(1)..rangesOfUseCase[4].elementAt(1),
//                    ifAction = { rowItem ->
//                        val cellidBigItem =
//                            sheetContainItem.getRow(rowItem)
//                                .getCell(columnNameToInt("h")).stringCellValue
//                        (cellidBigItem.length > 7 && cellidBigItem[6] == 'A' && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() % 4 == 1)
//                    },
//                    extraText = { _, _ ->
//                        ""
//                    })
//                processRows(rangeOfItem,
//                    rangesOfUseCase[4].elementAt(4)..rangesOfUseCase[4].elementAt(4),
//                    ifAction = { rowItem ->
//                        val cellidBigItem =
//                            sheetContainItem.getRow(rowItem)
//                                .getCell(columnNameToInt("h")).stringCellValue
//                        (cellidBigItem.length > 7 && cellidBigItem[6] == 'A' && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() % 4 == 3)
//                    },
//                    extraText = { _, _ ->
//                        ""
//                    })
//                processRows(rangeOfItem,
//                    rangesOfUseCase[4].elementAt(5)..rangesOfUseCase[4].elementAt(5),
//                    ifAction = { rowItem ->
//                        val cellidBigItem =
//                            sheetContainItem.getRow(rowItem)
//                                .getCell(columnNameToInt("h")).stringCellValue
//                        (cellidBigItem.length > 7 && cellidBigItem[6] == 'A' && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() % 4 == 0)
//                    },
//                    extraText = { _, _ ->
//                        ""
//                    })
//                processRows(rangeOfItem,
//                    rangesOfUseCase[4].elementAt(6)..rangesOfUseCase[4].elementAt(6),
//                    ifAction = { rowItem ->
//                        val cellidBigItem =
//                            sheetContainItem.getRow(rowItem)
//                                .getCell(columnNameToInt("h")).stringCellValue
//                        (cellidBigItem.length > 7 && cellidBigItem[6] == 'A' && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() % 4 == 3)
//                    },
//                    extraText = { _, _ ->
//                        ""
//                    })
//                processRows(rangeOfItem,
//                    rangesOfUseCase[4].elementAt(7)..rangesOfUseCase[4].elementAt(7),
//                    ifAction = { rowItem ->
//                        val cellidBigItem =
//                            sheetContainItem.getRow(rowItem)
//                                .getCell(columnNameToInt("h")).stringCellValue
//                        (cellidBigItem.length > 7 && cellidBigItem[6] == 'A' && "${cellidBigItem[7]}${cellidBigItem[8]}".toInt() % 4 == 0)
//                    },
//                    extraText = { _, _ ->
//                        ""
//                    })

                //back
                processRows(rangeOfItem, rangesOfUseCase[8],
                    ifAction = { rowItem ->

                        val cellidBigItem = sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("h")).stringCellValue

                        rowItem == rangeOfItem.last() && (cutAndFormatString(cellidBigItem, takeChars = 8) != "ST_SF_09")
                    },
                    extraText = { _, _ ->
                        ""
                    })
                ////////////////////////////
            }
            break
        }
        currentRow++
    }
}