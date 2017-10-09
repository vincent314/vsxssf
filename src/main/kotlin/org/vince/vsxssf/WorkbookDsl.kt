package org.vince.vsxssf

import java.io.File
import kotlin.coroutines.experimental.buildSequence

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

