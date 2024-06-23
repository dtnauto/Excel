package com.example.excel.excelfile

import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException


fun openExcelFile(filePath: String): XSSFWorkbook? {
    return try {
        FileInputStream(filePath).use { inputStream ->
            WorkbookFactory.create(inputStream) as XSSFWorkbook
        }
    } catch (e: IOException) {
        e.printStackTrace()
        println("Có lỗi xảy ra khi mở tệp Excel. Hãy đảm bảo rằng tệp không bị khóa bởi ứng dụng khác.")
        null
    }
}

fun openExcelFile(filePath: String, sheetName: String): XSSFSheet? {
    return try {
        FileInputStream(filePath).use { inputStream ->
            val workbook = WorkbookFactory.create(inputStream)
            val sheet = workbook.getSheet(sheetName)
            if (sheet == null) {
                println("Không tìm thấy sheet có tên \"$sheetName\" trong file Excel.")
                return null
            }
            sheet as XSSFSheet // Ép kiểu sang XSSFSheet và trả về
        }
    } catch (e: IOException) {
        e.printStackTrace()
        println("Có lỗi xảy ra khi mở tệp Excel. Hãy đảm bảo rằng tệp không bị khóa bởi ứng dụng khác.")
        null
    } catch (e: Exception) {
        e.printStackTrace()
        println("Đã xảy ra lỗi không xác định khi mở tệp Excel.")
        null
    }
}

fun openExcelFile(workbook: XSSFWorkbook, sheetName: String): XSSFSheet? {
    return try {
            val sheet = workbook.getSheet(sheetName)
            if (sheet == null) {
                println("Không tìm thấy sheet có tên \"$sheetName\" trong file Excel.")
                return null
            }
            sheet as XSSFSheet // Ép kiểu sang XSSFSheet và trả về
    } catch (e: IOException) {
        e.printStackTrace()
        println("Có lỗi xảy ra khi mở tệp Excel. Hãy đảm bảo rằng tệp không bị khóa bởi ứng dụng khác.")
        null
    } catch (e: Exception) {
        e.printStackTrace()
        println("Đã xảy ra lỗi không xác định khi mở tệp Excel.")
        null
    }
}

fun saveExcelFile(workbook: XSSFWorkbook, filePath: String) {
    try {
        FileOutputStream(filePath).use { fileOut ->
            workbook.write(fileOut)
        }
        workbook.close()
    } catch (e: IOException) {
        e.printStackTrace()
        println("Có lỗi xảy ra khi ghi tệp Excel. Hãy đảm bảo rằng tệp không bị khóa bởi ứng dụng khác.")
    }
}

fun saveExcelFile(sheet: XSSFSheet, filePath: String) {
    try {
        val workbook = sheet.workbook // Lấy workbook từ sheet
        FileOutputStream(filePath).use { fileOut ->
            workbook.write(fileOut)
        }
        workbook.close()
    } catch (e: IOException) {
        e.printStackTrace()
        println("Có lỗi xảy ra khi ghi tệp Excel. Hãy đảm bảo rằng tệp không bị khóa bởi ứng dụng khác.")
    }
}

fun columnNameToInt(columnName: String): Int {
    var result = 0
    var power = 1

    // Duyệt ngược lại chuỗi columnName từ phải sang trái
    for (i in columnName.length - 1 downTo 0) {
        val char = columnName[i].toUpperCase() // Chuyển đổi ký tự thành chữ hoa để dễ xử lý
        val value = char.toInt() - 'A'.toInt() // Chuyển đổi ký tự thành giá trị từ 0 (cho A) đến 25 (cho Z)
        result += value * power // Cộng dồn giá trị
        power *= 26 // Tăng mũ để tính cột tiếp theo
    }

    return result
}
