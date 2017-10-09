package org.vince.vsxssf

import org.apache.poi.hssf.util.HSSFColor

data class StyleOptions(
        val center: Boolean = false,
        val verticalCenter: Boolean = false,
        val bold: Boolean = false,
        val border: Boolean = false,
        val bgColor: HSSFColor.HSSFColorPredefined? = null,
        val wrapText: Boolean = false,
        val fontHeight: Short? = null
)
