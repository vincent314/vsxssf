package org.vince.vsxssf

import java.io.File
import java.io.OutputStream
import kotlin.coroutines.experimental.buildSequence

@Suppress("unused")
fun workbook(filename: String, init: Workbook.() -> Unit) {
    workbook(File(filename).outputStream(), init)
}

fun workbook(out: OutputStream, init: Workbook.() -> Unit){
    val workbook = Workbook()
    workbook.init()
    workbook.write(out)
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

