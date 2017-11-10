Very Simple XSSF
=============

VSXSSF is a wrapper library over [Apache POI-XSSF](https://poi.apache.org/spreadsheet/) written in Kotlin. It defines 
a DSL which simplify writing Excel files.

Usage example
--------

```kotlin
import org.vince.vsxssf.StyleOptions
import org.vince.vsxssf.workbook

val actors = arrayOf("William Hartnell", "Patrick Troughton", "Jon Pertwee", "Tom Baker", "Peter Davison", "Colin Baker", "Sylvester McCoy", "Paul McGann",
        "Christopher Eccleston", "David Tennant", "Matt Smith", "Peter Capaldi", "Jodie Whittaker")
workbook(filename = "/tmp/doctor-who.xlsx") {
    sheet(sheetName = "Actors", 
            colSize = arrayOf(2000, 4000)) {
        style(styleName = "title", 
                options = StyleOptions(bold = true))
        row {
            cell("Doctor #", "title")
            cell("Name", "title")
        }
        actors.forEachIndexed { index, name ->
            row {
                cell((index+1).toString())  // doctor's number
                cell(name)                  // actor's name
            }
        }
    }
}
```

Features
---------

* Creates Excel Workbook, Sheets, Rows and Cells
* Basic styling options (center, font size, colors)
* Page setup (page size and direction)
* Merged cells
* Repeated header
* Footer

Functions
----------

### workbook

Creates an Excel workbook file.

| attribute | type    | description                |
|-----------|---------|----------------------------|
| filename  | String  | destination .xlsx filename |

### sheet

_context: Workbook_

Creates an Excel sheet

| attribute     | type          | description                                 |
|---------------|---------------|---------------------------------------------|
| sheetName     | String        |sheet title                                  |
| colSize       | Array[Int]    |(madatory) column sizes                      |
| headerRows    | Int           |number of rows repeated on each printed page |
| landscape     | Boolean       |page layout in landscape                     |
| paperSize     | PaperSizeEnum |paper size (A4, LETTER)                      |

### style

_context: Workbook_

Register a new Style, associated with a name for further use

### StyleOptions

| attribute     | type              | description                               |
|---------------|-------------------|------------------------------------------ |
| center        | Boolean           |center text horizontally                   |
| verticalCenter| Boolean           |center text vertically                     |
| bold          | Boolean           |bold text                                  |
| border        | Boolean           |add a thin cell border                     |
| bgColor       | HSSFColorPredefined|background color                          |
| wrapText      | Boolean           |wrap text                                  |
| fontHeight    | Short             |set text font height                       |


### picture

_context: Workbook_

Register a new picture, associated with a name for further use

### footer

_context: Sheet_

Add footer texts

| attribute     | type             | description            |
|---------------|------------------|------------------------|
| left          | String           |left text               |
| center        | String           |center text             |
| right         | String           |right text              |

### row

_context: Sheet_

Add a new row

| attribute     | type            | description         |
|---------------|-----------------|---------------------|
| height        | Short           |define row height    |

### emptyRow

_context: Sheet_

Add an N empty new rows

| attribute     | type            | description         |
|---------------|-----------------|---------------------|
| height        | Short           |define row height    |
| repeat        | Int             |number of empty rows |

### cell

_context: Row_

Create a new Cell

| attribute     | type          | description           |
|---------------|---------------|-----------------------|
| text          | String        |cell content           |
| style         | String        |Style name as defined with style function |
| range         | Int           |merge N following cells|

### emptyCell

_context: Row_

Add an N empty new cells

| attribute     | type            | description         |
|---------------|-----------------|---------------------|
| style         | String           |set cell style       |
| repeat        | Int             |number of empty cells |


### picture

_context: Row_

Add a picture at current row

| attribute     | type            | description                 |
|---------------|-----------------|-----------------------------|
| col1          | Int             |set picture start from col1  |
| col2          | Int             |set picture end to col2      |
| pictureName   | String          |get picture from previously defined name      |
