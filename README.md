Very Simple XSSF
=============

VSXSSF is a wrapper library over [Apache POI-XSSF](https://poi.apache.org/spreadsheet/) written in Kotlin. It defines 
a DSL which simplify writing Excel files.

Usage example
--------

```kotlin
import org.vince.vsxssf.StyleOptions
import org.vince.vsxssf.workbook

val actors = arrayOf("William Hartnell", "Patrick Troughton", "Jon Pertwee", "Tom Baker", 
        "Peter Davison", "Colin Baker", "Sylvester McCoy", "Paul McGann",
        "Christopher Eccleston", "David Tennant", "Matt Smith", "Peter Capaldi", "Jodie Whittaker")
workbook("/tmp/doctor-who.xlsx") {
    sheet("Actors", arrayOf(2000, 4000)) {
        style("title", StyleOptions(bold = true)) // define some styles
        row {
            // render header
            cell("Doctor #", "title")
            cell("Name", "title")
        }
        actors.forEachIndexed { index, name ->
            row {
                cell((index+1).toString())  // doctor's number
                cell(name)                  // actor name
            }
        }
    }
}
```