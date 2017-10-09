package org.vince.vsxssf

import org.apache.poi.hssf.usermodel.HeaderFooter
import org.apache.poi.xssf.streaming.SXSSFSheet

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
