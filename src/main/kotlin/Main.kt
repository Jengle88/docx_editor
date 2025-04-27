package ru.jengle88

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xwpf.usermodel.*
import java.io.File
import java.io.FileInputStream
import java.text.SimpleDateFormat

fun main() {
//    val srcFile = File("docFile.docx")
//    val dstFile = File("docFile2.docx")
    println("Start parsing...")
    println("I'm in ${File("").canonicalPath}")
    println()

    val valuesFile = File("values.xlsx")
    val valuesTable = loadDataFromXlsx(valuesFile) ?: return

    for (_values in valuesTable) {
        val (currentFolder, table) = _values
        println("Start processing folder ${currentFolder}...")

        val srcFile = File(currentFolder, "шаблон.docx")

        val masksFile = File(currentFolder, "маски.txt")

        if (!srcFile.exists() || !masksFile.exists()) {
            println("Error: Folder $currentFolder not found!")
            continue
        }

        val masks = masksFile.readText().split(";").run {
            if (isNotEmpty() && last().isEmpty()) {
                dropLast(1)
            } else {
                this
            }
        }.map { mask -> mask.trim() }

        for ((index, row) in table.withIndex()) {
            val rowValues = row.take(masks.size).map { it.trim() }
            val docx = XWPFDocument(srcFile.inputStream())
            if (masks.size != rowValues.size) {
                println("Error: Masks and values must have same sizes!")
                continue
            }
            val fileNumber = index + 1
            var filename = "dstFile$fileNumber.docx"
            println("Start replacing $fileNumber file...")
            masks.zip(rowValues).forEach { (mask, value) ->
                if (mask == "\$filename\$") {
                    val fixedFilename = value
                        .replace("\\", "_")
                        .replace("/", "_")
                    filename = "${fixedFilename}.docx"
                }
                else if (mask != "") {
                    replaceTextInDocument(docx, mask, value)
                }
            }
            println("Finish replacing $fileNumber file.")
            val outputDir = File("out", currentFolder)
            if (!outputDir.exists()) {
                outputDir.mkdirs()
            }
            val dstFile = File(outputDir.path, filename)
            docx.write(dstFile.outputStream())
        }
        println("Finish processing folder ${currentFolder}.")
        println()
    }
//    println("Finish parsing. Press Enter to exit")
//    readln()
}

private fun loadDataFromXlsx(file: File): List<Pair<String, List<List<String>>>>? {
    val result = mutableListOf<Pair<String, List<List<String>>>>()
    val dateFormat = SimpleDateFormat("dd.MM.yyyy")

    if (!file.exists() || file.extension != "xlsx") {
        return null
    }
    FileInputStream(file).use { fis ->
        val workbook = XSSFWorkbook(fis)
        workbook.sheetIterator().forEach { sheet ->
            val table = mutableListOf<List<String>>()
            for (row in sheet) {
                val rowData = mutableListOf<String>()

                for ((index, cell) in row.withIndex()) {
                    val cellValue = when (cell.cellType) {
                        CellType.STRING -> cell.stringCellValue
                        CellType.NUMERIC -> {
                            if (DateUtil.isCellDateFormatted(cell)) {
                                dateFormat.format(cell.dateCellValue)
                            } else {
                                val numericValue = cell.numericCellValue
                                if (numericValue % 1 == 0.0) {
                                    numericValue.toInt().toString()
                                } else {
                                    String.format("%.2f", numericValue)
                                }
                            }
                        }
                        CellType.BOOLEAN -> cell.booleanCellValue.toString()
                        CellType.FORMULA -> {
                            try {
                                cell.stringCellValue
                            } catch (e: Exception) {
                                try {
                                    val numericValue = cell.numericCellValue
                                    if (numericValue % 1 == 0.0) {
                                        numericValue.toInt().toString()
                                    } else {
                                        String.format("%.2f", numericValue)
                                    }
                                } catch (e: Exception) {
                                    cell.cellFormula
                                }
                            }
                        }
                        CellType.BLANK -> ""
                        else -> ""
                    }

//                    rowData.add(cellValue)

                    // if you need to combine the last columns
                    if (index != row.toList().lastIndex) {
                        rowData.add(cellValue)
                    } else {
                        rowData[rowData.lastIndex] = rowData.last() + ' ' + cellValue
                    }
                }

                table.add(rowData)
            }
            result.add(sheet.sheetName to table.filter { it.any { s -> s.isNotBlank()} })
        }

        workbook.close()
    }
    return result
}


private fun replaceTextInDocument(document: XWPFDocument, oldText: String, newText: String) {
    diveToTablesAndReplace(document.tables, oldText, newText)
    replaceInParagraph(document.paragraphs, oldText, newText)
}

private fun diveToTablesAndReplace(tables: List<XWPFTable>, oldText: String, newText: String) {
    if (tables.isEmpty()) {
        return
    }
    for (table in tables) {
        diveToRowsAndReplace(table.rows, oldText, newText)
    }
}

private fun diveToRowsAndReplace(rows: List<XWPFTableRow>, oldText: String, newText: String) {
    if (rows.isEmpty()) {
        return
    }
    for (row in rows) {
        diveToCellsAndReplace(row.tableCells, oldText, newText)
    }
}

private fun diveToCellsAndReplace(cells: List<XWPFTableCell>, oldText: String, newText: String) {
    if (cells.isEmpty()) {
        return
    }
    for (cell in cells) {
        if (cell.tables.isNotEmpty()) {
            diveToTablesAndReplace(cell.tables, oldText, newText)
            continue
        }
        if (cell.paragraphs.isNotEmpty()) {
            replaceInParagraph(cell.paragraphs, oldText, newText)
        }
    }
}

private fun replaceInParagraph(paragraphs: List<XWPFParagraph>, oldText: String, newText: String) {
    for (paragraph in paragraphs) {
        var posInText: TextSegment? = paragraph.searchText(oldText, PositionInParagraph())
        while (posInText != null) {
            if (posInText.beginRun >= paragraph.runs.size || posInText.endRun >= paragraph.runs.size) {
                /* do nothing */
            } else if (posInText.beginRun == posInText.endRun) {
                val newText = paragraph.runs[posInText.beginRun].text().replace(oldText, newText)
                paragraph.runs[posInText.beginRun].setText(newText, 0)
            } else {
                val leftTextBeforeMask = paragraph.runs[posInText.beginRun].text()?.dropLastWhile { it != oldText.first() }?.dropLast(1)
                val rightTextAfterMask = paragraph.runs[posInText.endRun].text()?.dropWhile { it != oldText.last() }?.drop(1)

                for (i in posInText.endRun - 1 downTo posInText.beginRun + 1) {
                    paragraph.removeRun(i)
                }
                paragraph.runs[posInText.beginRun].setText(leftTextBeforeMask + newText, 0)
                paragraph.runs[posInText.beginRun + 1].setText(rightTextAfterMask, 0)
            }
            posInText = paragraph.searchText(oldText, PositionInParagraph())
        }
    }
}

private fun replaceInParagraphOld(paragraphs: List<XWPFParagraph>, oldText: String, newText: String) {
    for (paragraph in paragraphs) {
        val posInText = paragraph.searchText(oldText, PositionInParagraph()) ?: continue
        val leftTextBeforeMask = paragraph.runs[posInText.beginRun].text().takeWhile { it != oldText.first() }
        val rightTextAfterMask = paragraph.runs[posInText.endRun].text().takeLastWhile { it != oldText.last() }
        repeat(posInText.endRun - posInText.beginRun) {
            paragraph.removeRun(posInText.beginRun)
        }
        paragraph.runs[posInText.beginRun].apply {
            setText(leftTextBeforeMask + newText + rightTextAfterMask, 0)
        }
    }
}