package com.example.excel.process72

import com.example.excel.excelfile.columnNameToInt
import com.example.excel.excelfile.openExcelFile
import com.example.excel.excelfile.saveExcelFile
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File

// Hàm đọc file log và tìm từ khóa
fun findKeywordsInLog(logFilePath: String, keywords: List<String>): MutableList<String?> {
    // Khởi tạo một danh sách matchingLines có kích thước bằng danh sách keywords, với các phần tử ban đầu là null
    val matchingLines = MutableList<String?>(keywords.size) { null }

    // Đọc file log từ đường dẫn
    val logFile = File(file_Path + logFilePath.removePrefix("- ").trim())

    // Kiểm tra nếu file tồn tại
    if (logFile.exists()) {
        // Duyệt qua từng dòng trong file
        logFile.forEachLine { line ->
            // Tìm kiếm từ khóa trong dòng
            for ((index, keyword) in keywords.withIndex()) {
                if (line.contains(keyword)) {
                    // Bỏ qua số ký tự cố định ở đầu (ví dụ bỏ qua 34 ký tự đầu tiên)
                    val extractedLine = if (line.length > 34) line.substring(33) else "Line 0"
                    // Thêm dòng vào vị trí tương ứng với từ khóa trong danh sách keywords
                    matchingLines[index] = extractedLine
                    break // Dừng khi tìm thấy từ khóa đầu tiên để tránh thêm lại dòng cho các từ khóa khác
                }
            }
        }
    }
    return matchingLines
}



//Hàm xử lý dữ liệu từ file Excel trong khoảng từ dòng start đến dòng end
fun processExcelFile(workbook: XSSFWorkbook, sheetName: String, startRow: Int, endRow: Int) {
    val sheet = workbook.getSheet(sheetName)
//    Duyệt qua các hàng từ dòng startRow đến endRow(để ý startRow và endRow đã trừ 1 để khớp index 0)
    for (rowIndex in startRow..endRow) {

        val row = sheet.getRow(rowIndex) ?: continue

        val cellP = row.getCell(columnNameToInt("P")) //Cột P (16 = 15 theo index 0)
        val cellM = row.getCell(columnNameToInt("M")) //Cột M (13 = 12 theo index 0)
        val cellN = row.createCell(columnNameToInt("N")) //Cột N (14 = 13 theo index 0)

        val logFilePath = cellP?.stringCellValue ?: ""
        val keywordsBlock = cellM?.stringCellValue ?: "" //Tách từ khóa từ ô M

        val keywords = extractKeywords(keywordsBlock) //Tìm từ khóa trong file log
        val logMatches = findKeywordsInLog(logFilePath, keywords) //Format kết quả để điền vào ô N
        val result = formatLogMatches(logMatches) //Điền kết quả vào ô N
        cellN.setCellValue(result)

//        // Tạo định dạng cho ô (top align và align left)
//        val cellStyle = workbook.createCellStyle()
//        cellStyle.alignment = HorizontalAlignment.LEFT // Căn trái
//        cellStyle.verticalAlignment = VerticalAlignment.TOP // Căn trên
//        // Áp dụng định dạng cho ô
//        cellN.cellStyle = cellStyle
    }
}

fun extractKeywords(keywordsBlock: String): List<String> {
    // Tách thành các dòng
    return keywordsBlock.lines() // Chia văn bản thành từng dòng
        .filter { it.trim().startsWith("- ") } // Chỉ lấy các dòng bắt đầu bằng "- "
        .map { it.trim().removePrefix("- ").trim() } // Loại bỏ "- " và các khoảng trắng thừa
        .toList() // Chuyển thành danh sách
}


//Hàm định dạng kết quả cho cột N
fun formatLogMatches(logMatches: MutableList<String?>): String {
    val formattedMatches =
        logMatches.mapIndexed { index, line -> "(${index+1}) Confirm with log:\n- $line" }
    return formattedMatches.joinToString("\n")
}

const val file_Path =
//    "D:\\svn\\GAM.IVI.MCDC_IVI\\trunk\\01.Document\\01.EngineeringDocument\\Vehicle\\00.Output\\71_SWE4_ENG7.2\\07_安全装備設定\\Feature_1\\"
    "C:\\Users\\daotr\\Desktop\\New folder//"
const val file_Name =
    "VehicleApp-MCDC-SWE4-TSR_結合テスト_安全装備設定_F1_usecase_ST_SF_014-NhanDT53 - Copy.xlsx"

fun main() {
    val filePath = file_Path + file_Name

    val workbook = openExcelFile(filePath)

    if (workbook != null) {

//        val scanner = Scanner(System.`in`)
//        Nhập đường dẫn file Excel
//        println("Nhập đường dẫn file Excel: ")

//        val excelFilePath = scanner.nextLine()
//        Nhập dòng bắt đầu
//        println("Nhập số dòng bắt đầu (bao gồm): ")
//        val startRow = 8//scanner.nextInt()

//        Nhập dòng kết thúc
//        println("Nhập số dòng kết thúc (bao gồm): ")
//        val endRow = 9 //scanner.nextInt()

//        Xử lý file Excel từ dòng startRow đến endRow
        processExcelFile(workbook, "試験項目(ユースケース)", 7, 7)

//        ket thuc file
        saveExcelFile(workbook, filePath)
    }
}