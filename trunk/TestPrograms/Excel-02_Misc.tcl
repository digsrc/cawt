# Test miscellaneous CawtExcel procedures like setting colors, column width,
# inserting formulas and searching.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}
package require cawt

# Number of test rows and columns being generated.
set numRows  10
set numCols   3

# Generate row list with test data
for { set i 1 } { $i <= $numCols } { incr i } {
    lappend rowList $i
}

# Open Excel, show the application window and create a workbook.
set appId [Excel Open true]
set workbookId [Excel AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-02_Misc"]
append xlsFile [Excel GetExtString $appId]
file delete -force $xlsFile

# Select the first - already existing - worksheet,
# set its name and fill it with data.
set worksheetId [Excel GetWorksheetIdByIndex $workbookId 1]
Excel SetWorksheetName $worksheetId "ExcelMisc"

for { set row 1 } { $row <= $numRows } { incr row } {
    Excel SetRowValues $worksheetId $row $rowList
}

# Use different range selection procedures and test various
# formatting, color and border procedures.
set rangeId [Excel SelectCellByIndex $worksheetId 2 1 true]
Excel SetRangeFillColor $rangeId 255 0 0
Excel SetRangeTextColor $rangeId 0 255 0
Excel SetRangeBorders $rangeId
Cawt CheckList [list 255 0 0] [Excel GetRangeFillColor $rangeId] "Fill color of cell 2,1"

set rangeId [Excel SelectCellByIndex $worksheetId 3 1 true]
Excel SetRangeFillColor $rangeId 0 255 0
Excel SetRangeTextColor $rangeId 0 0 255
Cawt CheckList [list 0 255 0] [Excel GetRangeFillColor $rangeId] "Fill color of cell 3,1"

set rangeId [Excel SelectCellByIndex $worksheetId 4 1 true]
Excel SetRangeFillColor $rangeId 0 0 255
Excel SetRangeTextColor $rangeId 255 0 0
Excel SetRangeBorders $rangeId xlThick
Cawt CheckList [list 0 0 255] [Excel GetRangeFillColor $rangeId] "Fill color of cell 4,1"

set rangeId [Excel SelectRangeByIndex $worksheetId 5 1 5 1 true]
Excel SetRangeFillColor $rangeId 255 0 0
Excel SetRangeTextColor $rangeId 0 255 0

set rangeId [Excel SelectRangeByIndex $worksheetId 6 1 7 2 true]
Excel SetRangeFillColor $rangeId 0 255 0
Excel SetRangeTextColor $rangeId 0 0 255
Excel SetRangeBorders $rangeId xlThin xlDash
Cawt CheckList {6 1 7 2} [Excel GetRangeAsIndex $rangeId] "GetRangeAsIndex"

set rangeId [Excel SelectRangeByString $worksheetId "A8:C10" true]
Excel SetRangeFillColor $rangeId 0 0 255
Excel SetRangeTextColor $rangeId 255 0 0
Excel SetRangeFormat $rangeId "real" [Excel GetNumberFormat $appId "0" "000"]
Cawt CheckString "A8:C10" [Excel GetRangeAsString $rangeId] "GetRangeAsString"

# Test setting a formula.
set cell [Excel SelectCellByIndex $worksheetId 1 [expr $numCols + 2] true]
$cell Formula "=TODAY()"
Cawt CheckString "=TODAY()" [$cell Formula] "cellId Formula"
puts "  FormulaLocal: [$cell FormulaLocal]"

# Test the border capabilities.
Excel SetCellValue $worksheetId  3 [expr $numCols + 2] "Hallo"
Excel SetCellValue $worksheetId  4 [expr $numCols + 2] "Holla"

set rangeId [Excel SelectCellByIndex $worksheetId 3 [expr $numCols + 2] true]
Excel SetRangeBorders $rangeId xlThin xlContinuous 255 0 0

set rangeId [Excel SelectCellByIndex $worksheetId 4 [expr $numCols + 2] true]
Excel SetRangeBorders $rangeId xlThin xlContinuous 0 0 255

# Test merging cells.
Excel SetCellValue $worksheetId 1 6 "MergedCells"
set rangeId [Excel SelectRangeByIndex $worksheetId 1 6 2 8 true]
Excel SetRangeMergeCells $rangeId true
Excel SetRangeBorders $rangeId xlThick

# Test the search capabilities.
# Search only first 20 rows and columns for an existing string.
set str "Hallo"
set cell [Excel Search $worksheetId $str 1 1 20 20]
Cawt CheckNumber 2 [llength $cell] "Search $str"
if { [llength $cell] == 2 } {
    set rowNum [lindex $cell 0]
    set colNum [lindex $cell 1]
    Cawt CheckString "E3" "[Excel ColumnIntToChar $colNum]$rowNum" "Search $str"
}

# Search only first 20 rows and columns for a non-existing string.
set str "HalliHallo"
set cell [Excel Search $worksheetId $str 1 1 20 20]
Cawt CheckNumber 0 [llength $cell] "Search $str"

# Search whole worksheet for an existing string.
set str "Holla"
set cell [Excel Search $worksheetId $str]
Cawt CheckNumber 2 [llength $cell] "Search $str"
if { [llength $cell] == 2 } {
    set rowNum [lindex $cell 0]
    set colNum [lindex $cell 1]
    Cawt CheckString "E4" "[Excel ColumnIntToChar $colNum]$rowNum" "Search $str"
}

# Test different ways of setting column width.
# Set all used colums to fit, except columns 1 and 2.
Excel SetColumnsWidth $worksheetId 1 [expr $numCols + 6] 0
Excel SetColumnWidth $worksheetId 1 20
Excel SetColumnWidth $worksheetId 2 10

# Test copying a whole worksheet.
set copyWorksheetId [Excel AddWorksheet $workbookId "Copy"]
Excel CopyWorksheet $worksheetId $copyWorksheetId

Excel CopyWorksheetBefore $worksheetId $copyWorksheetId "CopyBefore"
Excel CopyWorksheetAfter  $worksheetId $copyWorksheetId "CopyAfter"

puts "Saving as Excel file: $xlsFile"
Excel SaveAs $workbookId $xlsFile "" false

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Excel Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
