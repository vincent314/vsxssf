package org.vince.vsxssf

import org.apache.poi.ss.usermodel.ClientAnchor
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.streaming.SXSSFCell
import org.apache.poi.xssf.streaming.SXSSFRow

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
