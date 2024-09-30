package com.example.excel.process72

import com.example.excel.excelfile.columnNameToInt
import com.example.excel.excelfile.openExcelFile
import com.example.excel.excelfile.saveExcelFile
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.InputStreamReader
import java.nio.charset.Charset


fun findKeywordsInLog(logFilePath: String, keywords: List<String>, addLine: Boolean = false): MutableList<String?> {
    // Khởi tạo một danh sách matchingLines có kích thước bằng danh sách keywords, với các phần tử ban đầu là null
    val matchingLines = MutableList<String?>(keywords.size) { null }

    // Đọc file log từ đường dẫn (bỏ ký tự thừa "- ")
    val logFile = File(file_Path + logFilePath.removePrefix("- ").trim())

    // Kiểm tra nếu file tồn tại và có đuôi .txt
    if (!logFile.exists()) {
        return matchingLines
    }

    if (!(logFilePath.endsWith(".txt", ignoreCase = true))) {
        return matchingLines
    }

    // Mở file với định dạng UTF-16LE
    logFile.inputStream().use { inputStream ->
        InputStreamReader(inputStream, Charset.forName("UTF-16LE")).use { reader ->
            var indexKeyword = 0 // Chỉ số từ khóa hiện tại
            var lineNumber = 1 // Số dòng hiện tại, bắt đầu từ 1

            // Đọc từng dòng trong file
            reader.forEachLine { line ->
                // Nếu đã duyệt hết tất cả từ khóa thì dừng lại
                if (indexKeyword >= keywords.size) return@forEachLine

                // Kiểm tra nếu từ khóa hiện tại có định dạng đúng với hasIndexFormat
                if (hasIndexFormat(keywords[indexKeyword])) {
                    matchingLines[indexKeyword] = keywords[indexKeyword]
                    // Nếu từ khóa có định dạng (??), bỏ qua và chuyển sang từ khóa tiếp theo
                    indexKeyword++
                    return@forEachLine
                }

                // Kiểm tra nếu dòng chứa từ khóa hiện tại
                if (line.contains(keywords[indexKeyword])) {
                    // Nếu addLine = true, thêm chỉ số dòng, nếu không thì trả về chuỗi trống ""
                    val linePrefix = if (addLine) "Line $lineNumber: " else ""

                    // Bỏ qua số ký tự cố định ở đầu (ví dụ bỏ qua 34 ký tự đầu tiên)
                    val extractedLine = if (line.length > 34) line.substring(33) else "Line 0"

                    // Thêm dòng vào vị trí tương ứng với từ khóa trong danh sách keywords
                    matchingLines[indexKeyword] = linePrefix + extractedLine

                    // Tăng chỉ số để chuyển sang từ khóa tiếp theo
                    indexKeyword++
                }

                // Tăng số dòng sau mỗi lần đọc
                lineNumber++
            }
        }
    }

    return matchingLines
}

//Hàm định dạng kết quả cho cột M
fun extractKeywords(keywordsBlock: String): List<String> {
    val regex = Regex("""\(.{1,2}\)""") // Regex tìm chuỗi có dạng (??) với ? là 1 hoặc 2 ký tự
    return keywordsBlock.lines() // Chia văn bản thành từng dòng
        .filter {
            it.trim().startsWith("- ") || regex.containsMatchIn(it)
        } // Lấy các dòng bắt đầu bằng "- " hoặc có chứa "(?)"
        .map {
            val trimmedLine = it.trim()
            if (regex.containsMatchIn(trimmedLine)) {
                regex.find(trimmedLine)?.value.orEmpty() // Trả về phần chuỗi có dạng "(?)"
            } else {
                trimmedLine.removePrefix("- ").trim() // Nếu không thì loại bỏ "- " và các khoảng trắng thừa
            }
        }.toList() // Chuyển thành danh sách
}

fun hasIndexFormat(input: String): Boolean {
    val regex = Regex("""^\(.{1,2}\)$""") // Định nghĩa chuỗi phải có định dạng "(?)" với ? là 1 hoặc 2 ký tự bất kỳ
    return regex.matches(input.trim()) // Kiểm tra xem chuỗi có khớp với định dạng hay không
}


//Hàm định dạng kết quả cho cột N
fun formatLogMatches(logMatches: MutableList<String?>): String {
    val formattedMatches = logMatches.mapIndexed { index, line ->
        if (line != null && hasIndexFormat(line)) {
            "$line Confirm with log:"
        } else {
            "- $line"
        }
    }
    return formattedMatches.joinToString("\n")
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
        val logMatches = findKeywordsInLog(logFilePath, keywords, true) //Format kết quả để điền vào ô N
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

const val file_Path =
//    "D:\\svn\\GAM.IVI.MCDC_IVI\\trunk\\01.Document\\01.EngineeringDocument\\Vehicle\\00.Output\\71_SWE4_ENG7.2\\07_安全装備設定\\Feature_2\\"
    "C:\\Users\\daotr\\Desktop\\New folder//"
const val file_Name =
    "VehicleApp-MCDC-SWE4-TSR_結合テスト_安全装備設定_F1_usecase_ST_SF_014-NhanDT53 - Copy.xlsx"
//    "VehicleApp-MCDC-SWE4-TSR_結合テスト_安全装備設定_F2_Usecase - Round2 - NhanDT53 - Copy.xlsm"

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
        processExcelFile(workbook, "試験項目(ユースケース)", 7, 7)//52

//        ket thuc file
        saveExcelFile(workbook, filePath)
    }
}