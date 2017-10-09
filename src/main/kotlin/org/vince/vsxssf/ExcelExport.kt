package org.vince.vsxssf

import org.apache.poi.hssf.usermodel.HeaderFooter
import org.apache.poi.hssf.util.HSSFColor
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.streaming.SXSSFCell
import org.apache.poi.xssf.streaming.SXSSFRow
import org.apache.poi.xssf.streaming.SXSSFSheet
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import java.io.File
import java.nio.file.Files
import java.nio.file.Paths
import kotlin.coroutines.experimental.buildSequence

@Suppress("unused")
class Workbook {
    val workbook: SXSSFWorkbook = SXSSFWorkbook()
    var styles: LinkedHashMap<String, CellStyle> = LinkedHashMap()
    var pictures: LinkedHashMap<String, Int> = LinkedHashMap()


    fun sheet(sheetName: String? = null,
              colSize: Array<Int>,
              headerRows: Int? = null,
              landscape: Boolean = false,
              init: Sheet.() -> Unit
    ): Sheet {
        val sxssfSheet = workbook.createSheet(sheetName)
                .also {
                    it.fitToPage = true
                    it.printSetup.fitWidth = 1
                    it.printSetup.fitHeight = 0
                    it.printSetup.landscape = landscape
                    if (headerRows != null) {
                        it.repeatingRows = CellRangeAddress(0, headerRows, 0, colSize.size)
                    }
                    colSize.forEachIndexed { idx, width ->
                        it.setColumnWidth(idx, width)
                    }
                }



        return Sheet(sxssfSheet, this)
                .also { it.init() }
    }

    fun write(file: File) {
        workbook.write(file.outputStream())
    }

    fun style(styleName: String, options: StyleOptions) {
        styles[styleName] = workbook.createCellStyle()
                .also {
                    it.wrapText = options.wrapText
                    if (options.bold) {
                        val font = workbook.createFont()
                        font.bold = true
                        it.setFont(font)
                    }
                    if (options.center) {
                        it.setAlignment(HorizontalAlignment.CENTER)
                    }
                    if (options.verticalCenter) {
                        it.setVerticalAlignment(VerticalAlignment.CENTER)
                    }

                    if (options.bgColor != null) {
                        it.setFillPattern(FillPatternType.SOLID_FOREGROUND)
                        it.fillForegroundColor = options.bgColor.color.index
                    }

                    if (options.border) {
                        it.setBorderTop(BorderStyle.THIN)
                        it.setBorderBottom(BorderStyle.THIN)
                        it.setBorderLeft(BorderStyle.THIN)
                        it.setBorderRight(BorderStyle.THIN)
                    }
                }
    }

    fun picture(pictureName: String, filename: String) {
        val bytes = Files.readAllBytes(Paths.get(filename))
        val logoPictureIndex = workbook.addPicture(bytes, org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_PNG)
        pictures[pictureName] = logoPictureIndex
    }
}

data class StyleOptions(
        val center: Boolean = false,
        val verticalCenter: Boolean = false,
        val bold: Boolean = false,
        val border: Boolean = false,
        val bgColor: HSSFColor.HSSFColorPredefined? = null,
        val wrapText: Boolean = false
)

@Suppress("unused")
class Sheet(private val sheet: SXSSFSheet, private val workbook: Workbook) {
    private val rowSequence = buildIterator()

    fun footer(left: String? = null, center: String? = null, right: String? = null) {
        with(sheet) {
            if (left != null) {
                footer.left = left
            }
            if (center != null) {
                footer.center = center
            }
            if (right != null) {
                footer.right = right
            }
        }
    }

    fun row(height: Short? = null, init: Row.() -> Unit): Row {
        val sxssfRow = sheet.createRow(rowSequence.next())
        if (height != null) {
            sxssfRow.height = height
        }

        return Row(sxssfRow, workbook)
                .also { it.init() }
    }

    fun emptyRow(height: Short? = null, repeat: Int = 1) {
        for (i in 1..repeat) {
            row(height) { }
        }
    }

    fun getPage(): String = HeaderFooter.page()

    fun getNumPage(): String = HeaderFooter.numPages()
}

@Suppress("unused")
class Row(private val row: SXSSFRow, private val workbook: Workbook) {
    private val cellSequence = buildIterator()
    private fun buildCell(style: String? = null): SXSSFCell =
            row.createCell(cellSequence.next())
                    .also {
                        if (style != null) {
                            it.cellStyle = workbook.styles[style]
                        }
                    }

    fun cell(text: String, style: String? = null, range: Int = 1) {
        with(buildCell(style)) {
            setCellValue(text)
            if (range > 1) {
                sheet.addMergedRegion(CellRangeAddress(rowIndex, rowIndex, columnIndex, columnIndex + range - 1))
                for (i in 1 until range) {
                    buildCell(style)
                }
            }
        }
    }

    fun emptyCell(style: String? = null, repeat: Int = 1) {
        for (i in 1..repeat) {
            buildCell(style)
        }
    }


    fun pageBreak() {
        row.sheet.setRowBreak(row.rowNum)
    }

    fun rowIndex(): Int = row.rowNum

    fun picture(col1: Int, col2: Int, pictureName:String) {
        val anchor = workbook.workbook.creationHelper.createClientAnchor()
                .also {
                    it.anchorType = ClientAnchor.AnchorType.MOVE_DONT_RESIZE
                    it.setCol1(col1)
                    it.setCol2(col2)
                    it.row1 = row.rowNum
                    it.row2 = row.rowNum
                }
        val patriarch = row.sheet.createDrawingPatriarch()

        if(workbook.pictures[pictureName] != null) {
            patriarch.createPicture(anchor, workbook.pictures[pictureName]!!)
                    .resize(1.0)
        }
    }
}

@Suppress("unused")
fun workbook(filename: String, init: Workbook.() -> Unit) {
    val workbook = Workbook()
    workbook.init()
    workbook.write(File(filename))
}

@Suppress("EXPERIMENTAL_FEATURE_WARNING")
fun buildIterator(): Iterator<Int> {
    return buildSequence {
        var value = 0
        while (true) {
            yield(value++)
        }
    }.iterator()
}

