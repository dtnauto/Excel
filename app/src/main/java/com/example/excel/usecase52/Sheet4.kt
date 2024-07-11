package com.example.excel.usecase52

import com.example.excel.excelfile.columnNameToInt
import com.example.excel.excelfile.openExcelFile
import com.example.excel.excelfile.saveExcelFile
import org.apache.poi.xssf.usermodel.XSSFWorkbook


fun main() {
    val filePath =
//        "C:\\Users\\daotr\\Desktop\\New Microsoft Excel Worksheet.xlsx"
        "D:\\working\\safety setting\\eng5.2\\New template.xlsx"

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
//        addRow(workbook, "Sheet4")

        // Gọi hàm để chèn các dòng mới
        copyRow(workbook, "Sheet1")

        // xu ly overView
        for (i in 500..workbook.getSheet("Sheet1").lastRowNum) {
            val currentRow = workbook.getSheet("Sheet1").getRow(i)
            val inputCell = currentRow?.getCell(columnNameToInt("k"))?.stringCellValue
            if (inputCell != null) {
                overView(inputCell)?.apply {
                    currentRow.createCell(columnNameToInt("n"))?.setCellValue(this)
                }
            }
        }


        // xu ly trigger
        for (i in 500..workbook.getSheet("Sheet1").lastRowNum) {
            val currentRow = workbook.getSheet("Sheet1").getRow(i)
            val inputCell = currentRow?.getCell(columnNameToInt("k"))?.stringCellValue
            if (inputCell != null) {
                trigger(inputCell)?.apply {
                    currentRow.createCell(columnNameToInt("m"))?.setCellValue(this)
                }
            }
        }


        // xu ly pre_condition
        for (i in 500..workbook.getSheet("Sheet1").lastRowNum) {
            val currentRow = workbook.getSheet("Sheet1").getRow(i)
            val inputCell = currentRow?.getCell(columnNameToInt("k"))?.stringCellValue
            if (inputCell != null) {
                pre_condition(inputCell)?.apply {
                    currentRow.createCell(columnNameToInt("i"))?.setCellValue(this)
                }
            }
        }

        // xu ly bigItem
        for (i in 500..workbook.getSheet("Sheet1").lastRowNum) {
            val currentRow = workbook.getSheet("Sheet1").getRow(i)
            val inputCell = currentRow?.getCell(columnNameToInt("k"))?.stringCellValue
            if (inputCell != null) {
                bigItem(inputCell)?.apply {
                    currentRow.createCell(columnNameToInt("e"))?.setCellValue(this)
                }
            }
        }

        // xu ly medium_item
        for (i in 500..workbook.getSheet("Sheet1").lastRowNum) {
            val currentRow = workbook.getSheet("Sheet1").getRow(i)
            val inputCell = currentRow?.getCell(columnNameToInt("k"))?.stringCellValue
            if (inputCell != null) {
                mediumItem(inputCell)?.apply {
                    currentRow.createCell(columnNameToInt("g"))?.setCellValue(this)
                }
            }
        }

        // xu ly category
        for (i in 500..workbook.getSheet("Sheet1").lastRowNum) {
            val currentRow = workbook.getSheet("Sheet1").getRow(i)
            val inputCell = currentRow?.getCell(columnNameToInt("k"))?.stringCellValue
            if (inputCell != null) {
                category(inputCell)?.apply {
                    currentRow.createCell(columnNameToInt("h"))?.setCellValue(this)
                }
            }
        }

//        for (i in 500..workbook.getSheet("Sheet1").lastRowNum) {
//            val currentRow = workbook.getSheet("Sheet1").getRow(i)
//            val inputCell = currentRow?.getCell(columnNameToInt("k"))?.stringCellValue
//            if (inputCell != null) {
//                cutAndFormatString(inputCell, endKeyword = "[bigItem]")?.apply {
//                    currentRow.createCell(columnNameToInt("k"))?.setCellValue(this.replace(" ()",""))
//                }
//            }
//        }

        // ket thuc file
        saveExcelFile(workbook, filePath)
    }
}

fun overView(inputCell: String?): String? {

    // Kiểm tra nếu inputCell là null hoặc chuỗi trống
    if (inputCell.isNullOrEmpty()) {
        return null
    }

    val extractedText = "<${cutAndFormatString(cutAndFormatString(inputCell, endKeyword = "[bigItem]"), "<", "]", 1)})".replace(" ()","")
    return when {
        inputCell.isEmpty() -> "kkk"
        inputCell.contains("Display value of") && inputCell.contains("setting item") && inputCell.contains(
            "Display item"
        ) -> {
            "- Get the setting value of $extractedText from VehicleAppService\n" +
                    "- Validate setting value and setting support\n" +
                    "- Display setting items according to the acquired setting values and support status"
        }

        inputCell.contains("Hide") && inputCell.contains("setting item") && inputCell.contains("Display item") -> {
            "- Get the setting value of $extractedText from VehicleAppService\n" +
                    "- Validate setting value and setting support\n"
        }

        inputCell.contains("Show") && inputCell.contains("setting item") && inputCell.contains("Display item") -> {
            "- Get the setting value of $extractedText from VehicleAppService\n" +
                    "- Validate setting value and setting support\n"
        }

        inputCell.contains("Tonedown") && inputCell.contains("setting item") && inputCell.contains("Display item") -> {
            "- Get the setting value of $extractedText from VehicleAppService\n" +
                    "- Validate setting value and setting support\n" +
                    "- Display setting items according to the acquired setting values and support status"
        }

        inputCell.contains("Toneup") && inputCell.contains("setting item") && inputCell.contains("Display item") -> {
            "- Get the setting value of $extractedText from VehicleAppService\n" +
                    "- Validate setting value and setting support\n" +
                    "- Display setting items according to the acquired setting values and support status"
        }

        inputCell.contains("Update the display value of") && inputCell.contains("setting item") && inputCell.contains(
            "Update item"
        ) -> {
            "- Received $extractedText change notification from VehicleAppService\n" +
                    "- Validate and update new setting value"
        }

        inputCell.contains("Hide") && inputCell.contains("setting item") && inputCell.contains("Update item") -> {
            "- Received $extractedText change notification from VehicleAppService\n" +
                    "- Validate setting support and update the tonedown/toneup status according setting support"
        }

        inputCell.contains("Show") && inputCell.contains("setting item") && inputCell.contains("Update item") -> {
            "- Received $extractedText change notification from VehicleAppService\n" +
                    "- Validate setting support and update the display (visible/invisible) according setting support"
        }

        inputCell.contains("Tonedown") && inputCell.contains("setting item") && inputCell.contains("Update item") -> {
            "- Received $extractedText change notification from VehicleAppService\n" +
                    "- Validate setting support and update the tonedown/toneup status according setting support"
        }

        inputCell.contains("Toneup") && inputCell.contains("setting item") && inputCell.contains("Update item") -> {
            "- Received $extractedText change notification from VehicleAppService\n" +
                    "- Validate setting support and update the display (visible/invisible) according setting support"
        }

        inputCell.contains("Change screen to") && inputCell.contains("by user operation") -> {
            "- Notify event change screen to StateManagement.\n" +
                    "- Display $extractedText screen\n" +
                    "- Make a preview display request to UXViewService"
        }

        inputCell.contains("Change setting of") && inputCell.contains("by user operation succeed") -> {
            "- Change the settings menu on the Safety Setting display content setting screen by user operation \n" +
                    "- Notify VehicleAPPService of a request to change $extractedText settings and service return result OK\n" +
                    "- Display the menu at the same time."
        }

        inputCell.contains("Change setting of") && inputCell.contains("by user operation failed") -> {
            "- Change the settings menu on the Safety Setting display content setting screen by user operation and service return result other than OK\n" +
                    "- Notify VehicleAPPService of a request to change $extractedText settings\n" +
                    "- Don't change setting status."
        }

        inputCell.contains("Click reset") && inputCell.contains("item") -> {
            "- Show $extractedText pop-up confirm reset"
        }

        inputCell.contains("Show") && inputCell.contains("pop-up") -> {
            "- Request OnScreenManager show $extractedText pop-up"
        }

        inputCell.contains("Pop-up") && inputCell.contains("is resetting") -> {
            "- Show $extractedText pop-up is resetting"
        }

        inputCell.contains("Pop-up") && inputCell.contains("cancelled") -> {
            "- Cancel $extractedText pop-up"
        }

        inputCell.contains("Pop-up") && inputCell.contains("success") -> {
            "- Show $extractedText Success Reset Dialog"
        }

        inputCell.contains("Pop-up") && inputCell.contains("failure") -> {
            "- Show $extractedText Fail Reset Dialog"
        }

        inputCell.contains("Success pop-up") && inputCell.contains("erasured") -> {
            "- Close $extractedText Success Reset Dialog"
        }

        inputCell.contains("Failure pop-up") && inputCell.contains("erasured") -> {
            "- Close $extractedText Fail Reset Dialog"
        }

        inputCell.contains("Back to") -> {
            "- Close ${bigItem(inputCell)} screen\n" +
                    "- Show $extractedText screen"
        }

        else -> null
    }
}

fun trigger(inputCell: String?): String? {
    // Kiểm tra nếu inputCell là null hoặc chuỗi trống
    if (inputCell.isNullOrEmpty()) {
        return null
    }

    val extractedText = "<${cutAndFormatString(cutAndFormatString(inputCell, endKeyword = "[bigItem]"), "<", "]", 1)})".replace(" ()","")

    return when {
        inputCell.isEmpty() -> "kkk"
        (inputCell.contains("Display [") || inputCell.contains("Display <")) && inputCell.contains("setting item") -> {
            "- When ${bigItem(inputCell)} displayed"
        }

        inputCell.contains("Hide") && inputCell.contains("setting item") -> {
            "- When ${bigItem(inputCell)} displayed"
        }

        inputCell.contains("Tonedown") && inputCell.contains("setting item") -> {
            "- When ${bigItem(inputCell)} displayed"
        }

        inputCell.contains("Update the display of") && inputCell.contains("setting item") -> {
            "- When receiving notification $extractedText setting Value change from VehicleAppService"
        }

        inputCell.contains("Update tonedown status of") && inputCell.contains("setting item") -> {
            "- When receiving notification $extractedText setting Value change from VehicleAppService"
        }

        inputCell.contains("Update display status of") && inputCell.contains("setting item") -> {
            "- When receiving notification $extractedText setting Value change from VehicleAppService"
        }

        inputCell.contains("Change screen to") && inputCell.contains("by user operation") -> {
            "- When user presses the $extractedText setting item on the ${bigItem(inputCell)} display content setting screen"
        }

        inputCell.contains("Change setting of") && inputCell.contains("by user operation succeed") -> {
            "- When user presses the $extractedText setting item on the ${bigItem(inputCell)} display content setting screen"
        }

        inputCell.contains("Change setting of") && inputCell.contains("by user operation failed") -> {
            "- When user presses the $extractedText setting item on the ${bigItem(inputCell)} display content setting screen"
        }

        inputCell.contains("Click reset") && inputCell.contains("item") -> {
            "- When user presses the $extractedText setting item on the ${bigItem(inputCell)} display content setting screen"
        }

        inputCell.contains("Click reset") && inputCell.contains("item") -> {
            "- When user presses the $extractedText setting item on the ${bigItem(inputCell)} display content setting screen"
        }

        inputCell.contains("Show") && inputCell.contains("pop-up") -> {
            "- When user press <${
                cutAndFormatString(
                    inputCell,
                    "<",
                    "リセット",
                    1
                )
            }リセット> [${
                cutAndFormatString(
                    inputCell,
                    "[",
                    " confirmation screen]",
                    1
                )
            }] button"
        }

        inputCell.contains("Pop-up") && inputCell.contains("cancelled") -> {
            "- When receiving operations from the user. User press [Cancel] on pop-up confirm reset"
        }

        inputCell.contains("Pop-up") && inputCell.contains("success") -> {
            "- When receiving change notification [Success Reset] from VehicleAppService"
        }

        inputCell.contains("Pop-up") && inputCell.contains("failure") -> {
            "- When receiving change notification [Fail Reset] from VehicleAppService"
        }

        (inputCell.contains("Success pop-up") || inputCell.contains("Failure pop-up")) && inputCell.contains(
            "erasured"
        ) -> {
            "- When 5 second passed"
        }

        inputCell.contains("Back to") -> {
            "- When user press [Back] button"
        }

        else -> null
    }
}

fun pre_condition(inputCell: String?): String? {
    // Kiểm tra nếu inputCell là null hoặc chuỗi trống
    if (inputCell.isNullOrEmpty()) {
        return null
    }

    val extractedText = "<${cutAndFormatString(cutAndFormatString(inputCell, endKeyword = "[bigItem]"), "<", "]", 1)})".replace(" ()","")

    return when {
        inputCell.isEmpty() -> "kkk"
        (inputCell.contains("Display [") || inputCell.contains("Display <")) && inputCell.contains("setting item") -> {
            "-"
        }

        inputCell.contains("Hide") && inputCell.contains("setting item") -> {
            "-"
        }

        inputCell.contains("Tonedown") && inputCell.contains("setting item") -> {
            "-"
        }

        inputCell.contains("Update the display of") && inputCell.contains("setting item") -> {
            "- Item ${bigItem(inputCell)} is displaying"
        }

        inputCell.contains("Update tonedown status of") && inputCell.contains("setting item") -> {
            "- Item ${bigItem(inputCell)} is displaying"
        }

        inputCell.contains("Update display status of") && inputCell.contains("setting item") -> {
            "- Item ${bigItem(inputCell)} is displaying"
        }

        inputCell.contains("Change screen to") && inputCell.contains("by user operation") -> {
            "- Item $extractedText is supported"
        }

        inputCell.contains("Change setting of") && inputCell.contains("by user operation succeed") -> {
            "- Item $extractedText is supported"
        }

        inputCell.contains("Change setting of") && inputCell.contains("by user operation failed") -> {
            "- Item $extractedText is supported"
        }

        inputCell.contains("Click reset") && inputCell.contains("item") -> {
            "- Item $extractedText is supported"
        }

        inputCell.contains("Show") && inputCell.contains("pop-up") -> {
            "- Item <${
                cutAndFormatString(
                    inputCell,
                    "<",
                    "リセット",
                    1
                )
            }リセット> [${
                cutAndFormatString(
                    inputCell,
                    "[",
                    " confirmation screen]",
                    1
                )
            }] is CLICKABLE state"
        }

        inputCell.contains("Pop-up") && inputCell.contains("is resetting") -> {
            "- Item $extractedText is resetting"
        }

        inputCell.contains("Pop-up") && inputCell.contains("cancelled") -> {
            "- Item $extractedText is resetting"
        }

        inputCell.contains("Pop-up") && inputCell.contains("success") -> {
            "-"
        }

        inputCell.contains("Pop-up") && inputCell.contains("failure") -> {
            "-"
        }

        (inputCell.contains("Success pop-up") || inputCell.contains("Failure pop-up")) && inputCell.contains(
            "erasured"
        ) -> {
            "-"
        }

        inputCell.contains("Back to") -> {
            "- Item $extractedText is supported"
        }

        else -> null
    }
}

fun bigItem(inputCell: String?): String? {
    // Kiểm tra nếu inputCell là null hoặc chuỗi trống
    if (inputCell.isNullOrEmpty()) {
        return null
    }

    var extractedText = cutAndFormatString(inputCell, "[bigItem]","[UseCase]", offsetLen = 9) ?: ""
    extractedText = "${cutAndFormatString(extractedText, "(", ")", 1)}\n" +
            "${cutAndFormatString(extractedText, endKeyword = " (")}]"

    return extractedText
}

fun mediumItem(inputCell: String?): String? {
    // Kiểm tra nếu inputCell là null hoặc chuỗi trống
    if (inputCell.isNullOrEmpty()) {
        return null
    }

    val extractedText = "<${cutAndFormatString(cutAndFormatString(inputCell, endKeyword = "[bigItem]"), "<", "]", 1)})".replace(" ()","")

    return when {
        inputCell.isEmpty() -> "kkk"
        (inputCell.contains("Display [") || inputCell.contains("Display <")) && inputCell.contains("setting item") -> {
            "- Display setting item"
        }

        inputCell.contains("Hide") && inputCell.contains("setting item") -> {
            "- Display setting item"
        }

        inputCell.contains("Tonedown") && inputCell.contains("setting item") -> {
            "- Display setting item"
        }

        inputCell.contains("Update the display of") && inputCell.contains("setting item") -> {
            "- Update display"
        }

        inputCell.contains("Update tonedown status of") && inputCell.contains("setting item") -> {
            "- Update display"
        }

        inputCell.contains("Update display status of") && inputCell.contains("setting item") -> {
            "- Update display"
        }

        inputCell.contains("Change screen to") && inputCell.contains("by user operation") -> {
            "- Change screen"
        }

        inputCell.contains("Change setting of") && inputCell.contains("by user operation succeed") -> {
            "- Change setting"
        }

        inputCell.contains("Change setting of") && inputCell.contains("by user operation failed") -> {
            "- Change setting"
        }

        inputCell.contains("Click reset") && inputCell.contains("item") -> {
            "- Reset"
        }

        inputCell.contains("Show") && inputCell.contains("pop-up") -> {
            "- Show popup"
        }

        inputCell.contains("Pop-up") && inputCell.contains("is resetting") -> {
            "- Press button on popup"
        }

        inputCell.contains("Pop-up") && inputCell.contains("cancelled") -> {
            "- Press button on popup"
        }

        inputCell.contains("Pop-up") && inputCell.contains("success") -> {
            "- Show popup"
        }

        inputCell.contains("Pop-up") && inputCell.contains("failure") -> {
            "- Show popup"
        }

        (inputCell.contains("Success pop-up") || inputCell.contains("Failure pop-up")) && inputCell.contains(
            "erasured"
        ) -> {
            "- Cancel popup"
        }

        inputCell.contains("Back to") -> {
            "- Back"
        }

        else -> null
    }
}

fun category(inputCell: String?): String? {
    // Kiểm tra nếu inputCell là null hoặc chuỗi trống
    if (inputCell.isNullOrEmpty()) {
        return null
    }

    val extractedText = "<${cutAndFormatString(cutAndFormatString(inputCell, endKeyword = "[bigItem]"), "<", "]", 1)})".replace(" ()","")
    val efs = "<${cutAndFormatString(cutAndFormatString(inputCell, startKeyword = "] (", endKeyword = "[bigItem]"), "(", ")")})"
    return when {
        inputCell.isEmpty() -> "kkk"
        (inputCell.contains("Display [") || inputCell.contains("Display <")) && inputCell.contains("setting item") -> {
            "- Normal"
        }

        inputCell.contains("Hide") && inputCell.contains("setting item") -> {
            "- Normal"
        }

        inputCell.contains("Tonedown") && inputCell.contains("setting item") -> {
            "- Normal"
        }

        inputCell.contains("Update the display of") && inputCell.contains("setting item") -> {
            "- Normal"
        }

        inputCell.contains("Update tonedown status of") && inputCell.contains("setting item") -> {
            "- Normal"
        }

        inputCell.contains("Update display status of") && inputCell.contains("setting item") -> {
            "- Normal"
        }

        inputCell.contains("Change screen to") && inputCell.contains("by user operation") -> {
            "- Normal"
        }

        inputCell.contains("Change setting of") && inputCell.contains("by user operation succeed") -> {
            "- Normal"
        }

        inputCell.contains("Change setting of") && inputCell.contains("by user operation failed") -> {
            "- Abnormal"
        }

        inputCell.contains("Click reset") && inputCell.contains("item") -> {
            "- Normal"
        }

        inputCell.contains("Show") && inputCell.contains("pop-up") -> {
            "- Normal"
        }

        inputCell.contains("Pop-up") && inputCell.contains("is resetting") -> {
            "- Normal"
        }

        inputCell.contains("Pop-up") && inputCell.contains("cancelled") -> {
            "- Normal"
        }

        inputCell.contains("Pop-up") && inputCell.contains("success") -> {
            "- Normal"
        }

        inputCell.contains("Pop-up") && inputCell.contains("failure") -> {
            "- Normal"
        }

        (inputCell.contains("Success pop-up") || inputCell.contains("Failure pop-up")) && inputCell.contains(
            "erasured"
        ) -> {
            "- Normal"
        }

        inputCell.contains("Back to") -> {
            "- Normal"
        }

        else -> null
    }
}

fun seq5dot2(inputCell: String?): String? {
    // Kiểm tra nếu inputCell là null hoặc chuỗi trống
    if (inputCell.isNullOrEmpty()) {
        return null
    }

//    val extractedText = "<${cutAndFormatString(cutAndFormatString(inputCell, endKeyword = "[bigItem]"), "<", ")", 1)})

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

        inputCell.contains("Pop-up") && inputCell.contains("failure") -> {
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

fun copyRow(workbook: XSSFWorkbook, sheetName: String) {

    val fileContainUseCase = "D:\\working\\safety setting\\eng5.2\\New template.xlsx"
//        openExcelFile("C:\\Users\\daotr\\Desktop\\New Microsoft Excel Worksheet.xlsx")!!
    val workbookContainUseCase = openExcelFile(fileContainUseCase)!!
    val sheetContainUseCase = workbookContainUseCase.getSheet("Sheet4")

    val sheet = workbook.getSheet(sheetName)

    var lastRowUseCase = sheetContainUseCase.lastRowNum
    var currentRowUseCase = 2 // fix current row
    while (currentRowUseCase <= lastRowUseCase) {
        val cellValue =
            sheetContainUseCase.getRow(currentRowUseCase)?.getCell(columnNameToInt("k"))?.toString()
                ?: ""
        if (cellValue.contains("screen1")) {
            break
        }
        currentRowUseCase++
    }

    var lastRow = sheet.lastRowNum
    var currentRow = 100 // fix current row
    while (currentRow <= lastRow) {
        val cellValue = sheet.getRow(currentRow)?.getCell(columnNameToInt("k"))?.toString() ?: ""
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

fun addRow(workbook: XSSFWorkbook, sheetName: String) {
    val fileContainItem = "D:\\working\\safety setting\\eng5.2\\New template.xlsx"
//        openExcelFile("C:\\Users\\daotr\\Desktop\\New Microsoft Excel Worksheet.xlsx")!!
    val workbookContainItem = openExcelFile(fileContainItem)!!
    val sheetContainItem = workbookContainItem.getSheet("Sheet3")

    val sheet = workbook.getSheet(sheetName)
    var lastRow = sheet.lastRowNum
    var currentRow = 2 // fix current row

    fun getRanges(
        startRow: Int = sheetContainItem.firstRowNum,
        endRow: Int = sheetContainItem.lastRowNum,
        columnGetRanges: Int,
    ): MutableList<IntRange> {
        val ranges = mutableListOf<IntRange>()

        var start = startRow
        var currentCellValue: String? = null

        for (row in startRow..if (endRow > sheetContainItem.lastRowNum) sheetContainItem.lastRowNum else endRow) {
            val cellValue =
                sheetContainItem.getRow(row).getCell(columnGetRanges).stringCellValue
            if (currentCellValue == null) {
                currentCellValue = cellValue
                start = row
            } else if (cellValue != currentCellValue) {
                ranges.add(start until row)
                currentCellValue = cellValue
                start = row
            }
        }
        ranges.add(start until (if (endRow > sheetContainItem.lastRowNum) sheetContainItem.lastRowNum else endRow) + 1)  // Add the last range

        return ranges
    }

    fun shiftAndCreateRow(newCellValue: String) {
        sheet.shiftRows(++currentRow, ++lastRow, 1)
        sheet.createRow(currentRow).createCell(columnNameToInt("k")).setCellValue(newCellValue)
    }

    fun createNewCellValue(
        rowUseCase: Int,
        rowItem: Int,
        extraText: String = ""
    ): String {
        var cellValueTemplate = sheetContainItem.getRow(rowUseCase)
            .getCell(columnNameToInt("h")).stringCellValue

        if (extraText == "[UseCase]Back") {
            val replacement =
                sheetContainItem.getRow(rowItem)
                    .getCell(columnNameToInt("b")).stringCellValue + " ()"

            cellValueTemplate = cellValueTemplate.replace("[]", replacement)
        } else if (extraText == "[UseCase]Reset" || extraText == "[UseCase]Change screen") {
            val replacement =
                "<${
                    cutAndFormatString(
                        sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("d")).stringCellValue,
                        "[", "]", 1
                    )
                }>" +
                        " [${
                            cutAndFormatString(
                                sheetContainItem.getRow(rowItem)
                                    .getCell(columnNameToInt("e")).stringCellValue,
                                "[", "]", 1
                            )
                        }]" +
                        " (${
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        })"

            cellValueTemplate = cellValueTemplate.replace("[]", replacement)
        } else {
            val replacement =
                "<${
                    cutAndFormatString(
                        sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("d")).stringCellValue,
                        "[", "]", 1
                    )
                }>" +
                        " [${
                            cutAndFormatString(
                                sheetContainItem.getRow(rowItem)
                                    .getCell(columnNameToInt("e")).stringCellValue,
                                "[", "]", 1
                            )
                        }]" + " ()"

            cellValueTemplate = cellValueTemplate.replace("[]", replacement)
        }
        val bigItem =
            sheetContainItem.getRow(rowItem).getCell(columnNameToInt("c")).stringCellValue

        return "$cellValueTemplate[bigItem]$bigItem$extraText"
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
        val cellValue = sheet.getRow(currentRow)?.getCell(columnNameToInt("k"))?.toString() ?: ""
        if (cellValue.contains("screen1")) {
//            val rangesOfItem = mutableListOf(2..14, 15..23, 24..27)
//            val rangesOfItem = mutableListOf(30..33,34..36,37..39,40..43,44..46)
//            val rangesOfItem = getRanges(112, 119, columnNameToInt("c"))
            val rangesOfItem = getRanges(120, 137, columnNameToInt("c"))
            println(rangesOfItem)
//            val rangesOfUseCase = mutableListOf(2..4, 5..7, 8..8, 9..10, 11..18, 19..19)
            val rangesOfUseCase = getRanges(2, 23, columnNameToInt("g"))
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
                    extraText = { rowItem, _ ->
                        "[UseCase]Display item"
                    })

                processRows(rangeOfItem, rangesOfUseCase[0].first + 1..rangesOfUseCase[0].last,
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 1)
                    },
                    extraText = { rowItem, _ ->
                        "[UseCase]Display item"
                    })

                //update
                processRows(rangeOfItem, rangesOfUseCase[1],
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] != 'A' || cellAction.length < 7)
                    },
                    extraText = { rowItem, _ ->
                        "[UseCase]Update item"
                    })

                processRows(rangeOfItem, rangesOfUseCase[1],
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 1)
                    },
                    extraText = { rowItem, _ ->
                        "[UseCase]Update item"
                    })

                // change
                processRows(rangeOfItem, rangesOfUseCase[2],
                    ifAction = { rowItem ->
                        val cellAction = sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == '0')
                    },
//                    extraText = { rowItem, _ ->
//                        sheetContainItem.getRow(rowItem)
//                            .getCell(columnNameToInt("f")).stringCellValue
//                    },
                    extraText = { rowItem, _ ->
                        "[UseCase]Change screen"
                    })

                processRows(rangeOfItem, rangesOfUseCase[3],
                    ifAction = { rowItem ->
                        val cellAction = sheetContainItem.getRow(rowItem)
                            .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.isEmpty())

                    },
                    extraText = { rowItem, _ ->
                        "[UseCase]Change setting"
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
                    extraText = { rowItem, _ ->
                        "[UseCase]Reset"
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(2)..rangesOfUseCase[4].elementAt(2),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 2)
                    },
                    extraText = { rowItem, _ ->
                        "[UseCase]Reset"
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(3)..rangesOfUseCase[4].elementAt(3),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 1)
                    },
                    extraText = { rowItem, _ ->
                        "[UseCase]Reset"
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(1)..rangesOfUseCase[4].elementAt(1),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 1)
                    },
                    extraText = { rowItem, _ ->
                        "[UseCase]Reset"
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(4)..rangesOfUseCase[4].elementAt(4),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 3)
                    },
                    extraText = { rowItem, _ ->
                        "[UseCase]Reset"
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(5)..rangesOfUseCase[4].elementAt(5),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 0)
                    },
                    extraText = { rowItem, _ ->
                        "[UseCase]Reset"
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(6)..rangesOfUseCase[4].elementAt(6),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 3)
                    },
                    extraText = { rowItem, _ ->
                        "[UseCase]Reset"
                    })
                processRows(rangeOfItem,
                    rangesOfUseCase[4].elementAt(7)..rangesOfUseCase[4].elementAt(7),
                    ifAction = { rowItem ->
                        val cellAction =
                            sheetContainItem.getRow(rowItem)
                                .getCell(columnNameToInt("f")).stringCellValue
                        (cellAction.length > 7 && cellAction[6] == 'A' && "${cellAction[7]}${cellAction[8]}".toInt() % 4 == 0)
                    },
                    extraText = { rowItem, _ ->
                        "[UseCase]Reset"
                    })

                //back
                processRows(rangeOfItem, rangesOfUseCase[5],
                    ifAction = { rowItem ->
                        rowItem == rangeOfItem.last()
                    },
                    { _, _ ->
                        "[UseCase]Back"
                    })
            }
        }
        currentRow++
    }
}