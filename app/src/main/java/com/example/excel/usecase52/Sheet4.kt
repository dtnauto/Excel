package com.example.excel.usecase52

import com.example.excel.excelfile.columnNameToInt
import org.apache.poi.xssf.usermodel.XSSFWorkbook

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
