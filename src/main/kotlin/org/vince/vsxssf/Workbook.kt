package org.vince.vsxssf

import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.vince.vsxssf.PaperSizeEnum.*
import java.io.OutputStream
import java.nio.file.Files
import java.nio.file.Paths

@Suppress("unused")
class Workbook {
    val workbook: SXSSFWorkbook = SXSSFWorkbook()
    var styles: LinkedHashMap<String, CellStyle> = LinkedHashMap()
    var pictures: LinkedHashMap<String, Int> = LinkedHashMap()

    fun sheet(sheetName: String? = null,
              colSize: Array<Int>,
              headerRows: Int? = null,
              landscape: Boolean = false,
              paperSize: PaperSizeEnum? = null,
              init: Sheet.() -> Unit
    ): Sheet {
        val sxssfSheet = workbook.createSheet(sheetName)
                .also {
                    it.fitToPage = true
                    it.printSetup.fitWidth = 1
                    it.printSetup.fitHeight = 0
                    it.printSetup.landscape = landscape
                    it.printSetup.paperSize = paperSize?.id ?: A4.id
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

    fun write(out:OutputStream) {
        workbook.write(out)
    }

    fun style(styleName: String, options: StyleOptions) {
        styles[styleName] = workbook.createCellStyle()
                .also {
                    it.wrapText = options.wrapText

                    val font = workbook.createFont()
                    font.bold = options.bold
                    if (options.fontHeight != null) {
                        font.fontHeightInPoints = options.fontHeight
                    }
                    it.setFont(font)

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
