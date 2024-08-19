# Changelog in Excel

`Worksheet_Change()` handler for automatically registrering metadata (author, date, time) in a changelog sheet.

These are the logical conditions that determine when rows are updated by the handler:

- __Cell value is changed (in column other than metadata columns)__
    - _Is this the only other column with a value?_
        - Add or update metadata
- __Cell value is removed (in column other than metadata columns)__
    - _Was this the only other column with a value?_
        - Remove metadata

_Updates to metadata are determined only by the first cell value added (or the last remaining) in a column other than metadata columns. The intent is to only capture metadata for a row when it is added to the changelog, by avoiding metadata updates when additional columns are added, changed, or corrected._

![Image of workbook using Worksheet_Change() handler](image-workbook.png)

## Requirements

- `.xlsm` file
- Macros enabled
- `Developer` tab added to ribbon

## Adding to VBAProject

These are the steps to add the code from `/Microsoft Excel Objects/Sheet.txt` in Excel.

1. Open the VBA Editor in `Developer > Visual Basic`
2. Open the code editor for the sheet in `VBAProject > Microsoft Excel Objects > Sheet`
    - `Sheet` will be the name you have given the sheet or a default sheet name
3. Paste in the code from `/Microsoft Excel Objects/Sheet.txt`
4. Close the VBA editor

## Configuration of metadata

These are the codelines and values that should be altered to customize the behavior of this handler.

| Description                                              | Code                               | Default value |
| -------------------------------------------------------- | ---------------------------------- | ------------- |
| Number of the column where author will be inserted       | `Const AuthorColumn As Integer = ` | `1`           |
| Number of the column where current date will be inserted | `Const DateColumn As Integer = `   | `2`           |
| Number of the column where time stamp will be inserted   | `Const TimeColumn As Integer = `   | `3`           |