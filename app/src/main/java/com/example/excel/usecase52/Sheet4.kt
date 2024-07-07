package com.example.excel.usecase52

import com.example.excel.excelfile.columnNameToInt
import com.example.excel.excelfile.openExcelFile
import com.example.excel.excelfile.saveExcelFile
import com.example.excel.screenhmi52.processColumnScreenSpec
import org.apache.poi.xssf.usermodel.XSSFWorkbook

fun main() {
    val filePath = "C:\\Users\\daotr\\Desktop\\New Microsoft Excel Worksheet.xlsx"
    val workbook = openExcelFile(filePath)
    if (workbook != null) {
//        addRow(workbook)
        addRow2(workbook)
        saveExcelFile(workbook, filePath)
    }
}

fun processColumnSheet4(workbook: XSSFWorkbook, sheetName: String, columnToProcess: Int) {
    val sheet = workbook.getSheet(sheetName)

    val firstRow = sheet.firstRowNum
    val lastRow = sheet.lastRowNum

    val rows = (firstRow..lastRow).toList().toIntArray()
    val columnSource = columnNameToInt("e")

    for (i in rows) {
        val currentRow = sheet.getRow(i)
        val inputCell = currentRow?.getCell(columnSource)?.stringCellValue
        if (inputCell != null) {
            when (columnToProcess) {
                5 -> {
                    overView(inputCell)?.apply {
                        currentRow.createCell(columnToProcess)?.setCellValue(this)
                    }
                }

                6 -> {
                    trigger(inputCell)?.apply {
                        currentRow.createCell(columnToProcess)?.setCellValue(this)
                    }
                }

                7 -> {
                    seq5dot2(inputCell)?.apply {
                        currentRow.createCell(columnToProcess)?.setCellValue(this)
                    }
                }
            }

        }
    }
}


fun overView(inputCell: String?): String? {

    // Kiểm tra nếu inputCell là null hoặc chuỗi trống
    if (inputCell.isNullOrEmpty()) {
        return null
    }

    val extractedText = cutAndFormatString(inputCell, "<", "]", 1)
    return when {
        inputCell.isEmpty() -> "kkk"
        (inputCell.contains("Display [") || inputCell.contains("Display <")) && inputCell.contains(
            "setting item"
        ) -> {
            "- Get the setting value of <$extractedText] from VehicleAppService\n" +
                    "- Display setting items according to the acquired setting values and support status"
        }

        inputCell.contains("Update the display of") && inputCell.contains("setting item") -> {
            "- Received <$extractedText] change notification from VehicleAppService\n" +
                    "- Get the value of <$extractedText] from VehicleAppService.\n" +
                    "- Update the <$extractedText] setting display follow value, support value from service, support value from subitem."
        }

        inputCell.contains("Change screen to") && inputCell.contains("by user operation") -> {
            "- Change to the <$extractedText] screen by user operation\n" +
                    "- Display the menu <$extractedText] and setting values according to the acquired support status\n" +
                    "- Obtain settings from VehicleAppService"
        }

        inputCell.contains("Change setting of") && inputCell.contains("by user operation") -> {
            "- Change the settings menu on the Safety Setting display content setting screen by user operation \n" +
                    "- Notify VehicleAPPService of a request to change <$extractedText] settings\n" +
                    "- Display the menu at the same time."
        }

        inputCell.contains("Click reset") && inputCell.contains("item") -> {
            "- Show <$extractedText] pop-up confirm reset"
        }

        inputCell.contains("Pop-up") && inputCell.contains("is resetting") -> {
            "- Show <$extractedText] pop-up is resetting"
        }

        inputCell.contains("Pop-up") && inputCell.contains("cancelled") -> {
            "- Cancel <$extractedText] pop-up"
        }

        inputCell.contains("Pop-up") && inputCell.contains("success") -> {
            "- Show <$extractedText] Success Reset Dialog"
        }

        inputCell.contains("Pop-up") && inputCell.contains("fail") -> {
            "- Show <$extractedText] Fail Reset Dialog"
        }

        inputCell.contains("Success pop-up") && inputCell.contains("erasured") -> {
            "- Close <$extractedText] Success Reset Dialog"
        }

        inputCell.contains("Failure pop-up") && inputCell.contains("erasured") -> {
            "- Close <$extractedText] Fail Reset Dialog"
        }

        inputCell.contains("Back to") -> {
            "- Display <$extractedText] screen"
        }

        else -> null
    }
}

fun trigger(inputCell: String?): String? {
    // Kiểm tra nếu inputCell là null hoặc chuỗi trống
    if (inputCell.isNullOrEmpty()) {
        return null
    }

    val extractedText = cutAndFormatString(inputCell, "<", "]", 1)

    return when {
        inputCell.isEmpty() -> "kkk"
        (inputCell.contains("Display [") || inputCell.contains("Display <")) && inputCell.contains(
            "setting item"
        ) -> {
            "- When displaying the <$extractedText] display content setting screen by user operation"
        }

        inputCell.contains("Update the display of") && inputCell.contains("setting item") -> {
            "- When receiving notification <$extractedText] setting Value change from VehicleAppService"
        }

        inputCell.contains("Change screen to") && inputCell.contains("by user operation") -> {
            "- When user presses the <$extractedText] setting item on the Safety Setting display content setting screen"
        }

        inputCell.contains("Change setting of") && inputCell.contains("by user operation") -> {
            "- When receiving operations from the user to change setting on the <$extractedText] setting screen"
        }

        inputCell.contains("Click reset") && inputCell.contains("item") -> {
            "- When receiving operations from the user. User press reset <$extractedText] item on setting menu"
        }

        inputCell.contains("Pop-up") && inputCell.contains("is resetting") -> {
            "- When receiving operations from the user. User press [Confirm] on pop-up confirm reset"
        }

        inputCell.contains("Pop-up") && inputCell.contains("cancelled") -> {
            "- When receiving operations from the user. User press [Cancel] on pop-up confirm reset"
        }

        inputCell.contains("Pop-up") && inputCell.contains("success") -> {
            "- When receiving change notification [Success Reset] from VehicleAppService"
        }

        inputCell.contains("Pop-up") && inputCell.contains("fail") -> {
            "- When receiving change notification [Fail Reset] from VehicleAppService"
        }

        (inputCell.contains("Success pop-up") || inputCell.contains("Failure pop-up")) && inputCell.contains(
            "erasured"
        ) -> {
            "- When 5 second passed"
        }

        inputCell.contains("Back to") -> {
            "- When receiving operations from the user. User presses [Back] button on Menu bar"
        }

        else -> null
    }
}

fun seq5dot2(inputCell: String?): String? {
    // Kiểm tra nếu inputCell là null hoặc chuỗi trống
    if (inputCell.isNullOrEmpty()) {
        return null
    }

//    val extractedText = cutAndFormatString(inputCell, "<", "]", 1)

    return when {
        inputCell.isEmpty() -> "kkk"
        (inputCell.contains("Display [") || inputCell.contains("Display <")) && inputCell.contains(
            "setting item"
        ) -> {
            "SQ-010. Display view"
        }

        inputCell.contains("Update the display of") && inputCell.contains("setting item") -> {
            "SQ-004 Update the setting display triggered by service"
        }

        inputCell.contains("Change screen to") && inputCell.contains("by user operation") -> {
            "SQ-003 Change screen"
        }

        inputCell.contains("Change setting of") && inputCell.contains("by user operation") -> {
            "SQ-005 User change setting by tapping switch button/radio button"
        }

        inputCell.contains("Click reset") && inputCell.contains("item") -> {
            "SQ-007. Reset setting TBD"
        }

        inputCell.contains("Pop-up") && inputCell.contains("is resetting") -> {
            "SQ-011.2. Show In Progress ONS"
        }

        inputCell.contains("Pop-up") && inputCell.contains("cancelled") -> {
            "SQ-011.5. Cancel ONS"
        }

        inputCell.contains("Pop-up") && inputCell.contains("success") -> {
            "SQ-011.3. Show Success ONS"
        }

        inputCell.contains("Pop-up") && inputCell.contains("fail") -> {
            "SQ-011.4. Show Failed ONS"
        }

        inputCell.contains("Success pop-up") && inputCell.contains("erasured") -> {
            "SQ-011.3. Show Success ONS"
        }

        inputCell.contains("Failure pop-up") && inputCell.contains("erasured") -> {
            "SQ-011.4. Show Failed ONS"
        }

        inputCell.contains("Back to") -> {
            "SQ-006. Back screen"
        }

        inputCell.contains("scroll") || inputCell.contains("Scroll") -> {
            "Using CommonUI's default behavior, the app doesn't have to handle it"
        }

        inputCell.contains("Start ") && inputCell.contains("app") -> {
            "SQ-001 Start Application"
        }

        inputCell.contains("Exitting ") -> {
            "SQ-008. Exit application"
        }

        inputCell.contains("Termiate ") -> {
            "SQ-008. Exit application"
        }

        inputCell.contains("Abnormal ") -> {
            "SQ-009. VehicleAppService abnormal termination\nSQ-001 Start Application"
        }

        inputCell.contains("PCS(") && cutAndFormatString(
            inputCell,
            ")",
            "PCS(ON)",
            1
        ) == "" -> {
            "SQ-004 Update the setting display triggered by service\n" +
                    "SQ-002 Display screen\n" +
                    "SQ-010. Display view\n" +
                    "SQ-008. Exit application"
        }

        inputCell.contains(cutAndFormatString(inputCell, ")", "PCS(ON)", 1) + "PCS(ON)") -> {
            "SQ-004 Update the setting display triggered by service\n" +
                    "SQ-002 Display screen\n" +
                    "SQ-010. Display view"
        }

        else -> null
    }
}

fun addRow(workbook: XSSFWorkbook) {
    val workbookContainItem =
        openExcelFile("C:\\Users\\daotr\\Desktop\\New Microsoft Excel Worksheet.xlsx")!!
    val sheetContainItem = workbookContainItem.getSheet("Sheet2")
    val sheet = workbook.getSheet("Sheet4")

    var lastRow = sheet.lastRowNum
    var currentRow = 2 // fix current row

    while (currentRow <= lastRow) { // kiểm tra điều kiện để điền vào ô
        val cellValue = sheet.getRow(currentRow)?.getCell(columnNameToInt("d"))?.toString() ?: ""

        // Kiểm tra điều kiện để chèn dòng mới
        if (cellValue.contains("ahihi")) {

            for (rowItem in 2..14) {
                // fix gia tri dien vao
                for (rowUseCase in 2..4) {
                    val newCellValue = sheetContainItem.getRow(rowUseCase)
                        .getCell(columnNameToInt("h")).stringCellValue
                        .replace(
                            "[]",
                            "<${
                                sheetContainItem.getRow(rowItem)
                                    .getCell(columnNameToInt("d")).stringCellValue
                            }>" +
                                    " [${
                                        sheetContainItem.getRow(rowItem)
                                            .getCell(columnNameToInt("e")).stringCellValue
                                    }]"
                        )
                    sheet.shiftRows(++currentRow, ++lastRow, 1)
                    sheet.createRow(currentRow).createCell(columnNameToInt("d"))
                        .setCellValue(newCellValue)
                }
            }

            for (rowItem in 2..14) {
                // fix gia tri dien vao
                for (rowUseCase in 5..7) {
                    val newCellValue = sheetContainItem.getRow(rowUseCase)
                        .getCell(columnNameToInt("h")).stringCellValue
                        .replace(
                            "[]",
                            "<${
                                sheetContainItem.getRow(rowItem)
                                    .getCell(columnNameToInt("d")).stringCellValue
                            }>" +
                                    " [${
                                        sheetContainItem.getRow(rowItem)
                                            .getCell(columnNameToInt("e")).stringCellValue
                                    }]"
                        )
                    sheet.shiftRows(++currentRow, ++lastRow, 1)
                    sheet.createRow(currentRow).createCell(columnNameToInt("d"))
                        .setCellValue(newCellValue)
                }
            }

            for (rowItem in 2..14) {
                val cellAction =
                    sheetContainItem.getRow(rowItem).getCell(columnNameToInt("f")).stringCellValue
                if (cellAction != null && cellAction.length > 7) {
                    if (cellAction[6] == '0') {
                        // fix gia tri dien vao
                        for (rowUseCase in 8..8) {
                            val newCellValue = sheetContainItem.getRow(rowUseCase)
                                .getCell(columnNameToInt("h")).stringCellValue
                                .replace(
                                    "[]",
                                    "<${
                                        sheetContainItem.getRow(rowItem)
                                            .getCell(columnNameToInt("d")).stringCellValue
                                    }>" +
                                            " [${
                                                sheetContainItem.getRow(rowItem)
                                                    .getCell(columnNameToInt("e")).stringCellValue
                                            }" +
                                            " ($cellAction)]"
                                )

                            sheet.shiftRows(++currentRow, ++lastRow, 1)
                            sheet.createRow(currentRow).createCell(columnNameToInt("d"))
                                .setCellValue(newCellValue)
                        }
                    }
                }
            }

            for (rowItem in 2..14) {
                val cellAction =
                    sheetContainItem.getRow(rowItem).getCell(columnNameToInt("f")).stringCellValue
                if (cellAction.isNullOrEmpty()) {
                    // fix gia tri dien vao
                    for (rowUseCase in 9..10) {
                        val newCellValue = sheetContainItem.getRow(rowUseCase)
                            .getCell(columnNameToInt("h")).stringCellValue
                            .replace(
                                "[]",
                                "<${
                                    sheetContainItem.getRow(rowItem)
                                        .getCell(columnNameToInt("d")).stringCellValue
                                }>" +
                                        " [${
                                            sheetContainItem.getRow(rowItem)
                                                .getCell(columnNameToInt("e")).stringCellValue
                                        }]"
                            )

                        sheet.shiftRows(++currentRow, ++lastRow, 1)
                        sheet.createRow(currentRow).createCell(columnNameToInt("d"))
                            .setCellValue(newCellValue)
                    }
                }
            }

            for (rowItem in 2..14) {
                val cellAction =
                    sheetContainItem.getRow(rowItem).getCell(columnNameToInt("f")).stringCellValue
                if (cellAction != null && cellAction.length > 7) {
                    if (cellAction[6] == 'A') {
                        // fix gia tri dien vao
                        for (rowUseCase in 11..17) {
                            val newCellValue = sheetContainItem.getRow(rowUseCase)
                                .getCell(columnNameToInt("h")).stringCellValue
                                .replace(
                                    "[]",
                                    "<${
                                        sheetContainItem.getRow(rowItem)
                                            .getCell(columnNameToInt("d")).stringCellValue
                                    }>" +
                                            " [${
                                                sheetContainItem.getRow(rowItem)
                                                    .getCell(columnNameToInt("e")).stringCellValue
                                            }" +
                                            " (${
                                                cellAction.replace(
                                                    "${cellAction[7]}${cellAction[8]}",
                                                    ("${cellAction[7]}${cellAction[8]}".toInt() + sheetContainItem.getRow(
                                                        rowUseCase
                                                    )
                                                        .getCell(columnNameToInt("i")).numericCellValue.toInt()).toString()
                                                        .padStart(2, '0')
                                                )
                                            })]"
                                )

                            sheet.shiftRows(++currentRow, ++lastRow, 1)
                            sheet.createRow(currentRow).createCell(columnNameToInt("d"))
                                .setCellValue(newCellValue)
                        }
                    }
                }
            }

            for (rowItem in 2..14) {
                if (rowItem == 14) {
                    // fix gia tri dien vao
                    for (rowUseCase in 18..18) {
                        val newCellValue = sheetContainItem.getRow(rowUseCase)
                            .getCell(columnNameToInt("h")).stringCellValue
                            .replace(
                                "[]",
                                "[mediumItem]"
                            )
                        sheet.shiftRows(++currentRow, ++lastRow, 1)
                        sheet.createRow(currentRow).createCell(columnNameToInt("d"))
                            .setCellValue(newCellValue)
                    }
                }
            }
        }
        currentRow++
    }
}

fun addRow2(workbook: XSSFWorkbook) {
    val workbookContainItem =
        openExcelFile("C:\\Users\\daotr\\Desktop\\New Microsoft Excel Worksheet.xlsx")!!
    val sheetContainItem = workbookContainItem.getSheet("Sheet2")
    val sheet = workbook.getSheet("Sheet4")

    var lastRow = sheet.lastRowNum
    var currentRow = 2 // fix current row

    fun shiftAndCreateRow(newCellValue: String) {
        sheet.shiftRows(++currentRow, ++lastRow, 1)
        sheet.createRow(currentRow).createCell(columnNameToInt("d")).setCellValue(newCellValue)
    }

    fun createNewCellValue(
        rowUseCase: Int,
        rowItem: Int,
        extraText: String = ""
    ): String {
        val cellValueTemplate = sheetContainItem.getRow(rowUseCase)
            .getCell(columnNameToInt("h")).stringCellValue

        return if (extraText == "[mediumItem]") {
            cellValueTemplate.replace("[]", "[mediumItem]")
        } else {
            val replacement = "<${
                sheetContainItem.getRow(rowItem).getCell(columnNameToInt("d")).stringCellValue
            }>" +
                    " [${
                        sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("e")).stringCellValue
                    }$extraText]"

            cellValueTemplate.replace("[]", replacement)
        }
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
                    shiftAndCreateRow(
                        createNewCellValue(
                            rowUseCase,
                            rowItem,
                            extraText(rowItem, rowUseCase)
                        )
                    )
                }
            }
        }
    }

    while (currentRow <= lastRow) {
        val cellValue = sheet.getRow(currentRow)?.getCell(columnNameToInt("d"))?.toString() ?: ""
        if (cellValue.contains("ahihi")) {

            processRows(2..14, 2..4,
                ifAction = { _ ->
                    true
                })
            processRows(2..14, 5..7,
                ifAction = { _ ->
                    true
                })

            processRows(2..14, 8..8,
                ifAction = { rowItem ->
                    val cellAction = sheetContainItem.getRow(rowItem)
                        .getCell(columnNameToInt("f")).stringCellValue
                    (cellAction.length > 7 && cellAction[6] == '0')
                },
                extraText = { rowItem, _ ->
                    val cellAction = sheetContainItem.getRow(rowItem)
                        .getCell(columnNameToInt("f")).stringCellValue

                    " ($cellAction)"
                })

            processRows(2..14, 9..10,
                ifAction = { rowItem ->
                    val cellAction = sheetContainItem.getRow(rowItem)
                        .getCell(columnNameToInt("f")).stringCellValue
                    (cellAction.isEmpty())
                })

            processRows(2..14, 11..17,
                ifAction = { rowItem ->
                    val cellAction =
                        sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("f")).stringCellValue
                    (cellAction.length > 7 && cellAction[6] == 'A')
                },
                extraText = { rowItem, rowUseCase ->
                    val cellAction = sheetContainItem.getRow(rowItem)
                        .getCell(columnNameToInt("f")).stringCellValue

                    val adjustedAction = cellAction.replace(
                        "${cellAction[7]}${cellAction[8]}",
                        ("${cellAction[7]}${cellAction[8]}".toInt() +
                                sheetContainItem.getRow(rowUseCase)
                                    .getCell(columnNameToInt("i")).numericCellValue.toInt()
                                ).toString().padStart(2, '0')
                    )

                    " ($adjustedAction)"
                })

            processRows(2..14, 18..18,
                ifAction = { rowItem ->
                    rowItem == 14
                },
                { _, _ ->
                    "[mediumItem]"
                })
        }
        currentRow++
    }
}


fun newRowIf(
    workbook: XSSFWorkbook,
    sheetName: String,
    ranges: Array<Int>,
    insertItems: Array<String>,
    insertValues: Array<String>
) {
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
        if (cellValue.contains("ahihi")) {//&& cellValue.contains("] screen")) {
            var newRow = currentRow + 1

            // Chèn dòng mới với các giá trị từ mảng insertItems và insertValues
            for (item in insertItems) {
                for (value in insertValues) {
                    if (item.isNotBlank()) {
                        sheet.shiftRows(newRow, lastRow + 1, 1)
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

