package org.vince.vsxssf

import org.apache.poi.hssf.usermodel.HSSFPrintSetup

enum class PaperSizeEnum(val id:Short) {
    A4(HSSFPrintSetup.A4_PAPERSIZE),
    LETTER(HSSFPrintSetup.LETTER_PAPERSIZE)
}