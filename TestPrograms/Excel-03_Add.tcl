# Test CawtExcel procedures for adding and deleting workbooks and worksheets.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

# Name of Excel file being generated.
# Number of rows, columns and worksheets being generated.
set xlsFile [file join [pwd] "testOut" "Excel-03_Add"]
set numRows   20
set numCols   10
set numSheets  3

# Generate row list with test data
for { set i 1 } { $i <= $numCols } { incr i } {
    lappend rowList "E-$i"
}

# Generate worksheets and fill them with the test data.
# We use the Open procedure, which should re-use an already
# open Excel instance.
for { set s 1 } { $s <= $numSheets } { incr s } {
    set appId [Excel Open]
    set workbookId [Excel GetActiveWorkbook $appId]
    if { ! [Cawt IsValidId $workbookId] } {
        set workbookId [Excel AddWorkbook $appId]
    }

    set worksheetIds($s) [Excel AddWorksheet $workbookId "Sheet-$s"]
    for { set i 1 } { $i <= $numRows } { incr i } {
        Excel SetRowValues $worksheetIds($s) $i $rowList
    }
}
Cawt CheckNumber 20 [Excel GetNumUsedRows $worksheetIds(2)] \
                       "Number of used rows in Sheet-2"
Cawt CheckNumber 10 [Excel GetNumUsedColumns $worksheetIds(2)] \
                       "Number of used columns in Sheet-2"

# Add another worksheet and fill it with header lines.
set worksheetId [Excel AddWorksheet $workbookId "HeaderRows"]
set startColumn 1
for { set i 1 } { $i <= $numRows } { incr i } {
    Excel SetHeaderRow $worksheetId $rowList $i $startColumn
    incr startColumn
}
set wsName [Excel GetWorksheetName $worksheetId]
Cawt CheckNumber 20 [Excel GetNumUsedRows $worksheetId] \
                       "Number of used rows in $wsName"
Cawt CheckNumber 29 [Excel GetNumUsedColumns $worksheetId] \
                       "Number of used columns in $wsName"

# Test retrieving parts of a row or column.
set rowValuesList [Excel GetRowValues $worksheetId 1 5]
Cawt CheckList [lrange $rowList 4 end] $rowValuesList "Values of row 1 (starting at column 5)"
Cawt CheckNumber 6 [llength $rowValuesList] "Number of retrieved row elements in $wsName"

set colValuesList [Excel GetColumnValues $worksheetId 7 2]
set colList [list E-6 E-5 E-4 E-3 E-2 E-1]
Cawt CheckList $colList $colValuesList "Values of column 7 (starting at row 2)"
Cawt CheckNumber 6 [llength $colValuesList] "Number of retrieved column elements in $wsName"

# Test different ways to delete a worksheet.
set num [Excel GetNumWorksheets $workbookId]
Cawt CheckNumber 5 $num "Number of worksheets before deletion of Sheet-1"
after 500

set sheetId [Excel GetWorksheetIdByName $workbookId "Sheet-1"]
Excel DeleteWorksheet $workbookId $sheetId

set num [Excel GetNumWorksheets $workbookId]
Cawt CheckNumber 4 $num "Number of worksheets before deletion of last sheet"
after 500

Excel DeleteWorksheetByIndex $workbookId $num

set num [Excel GetNumWorksheets $workbookId]
Cawt CheckNumber 3 $num "Number of worksheets finally"

# Append the default Excel filename extension.
append xlsFile [Excel GetExtString $appId]

# # Delete Excel file from previous test run.
file mkdir testOut
catch { file delete -force $xlsFile }

puts "Saving as Excel file: $xlsFile"
Excel SaveAs $workbookId $xlsFile

# Reopen the generated workbook in a new Excel instance.
set appId2 [Excel OpenNew]
set workbookId [Excel OpenWorkbook $appId2 $xlsFile]

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Excel Quit $appId
    Excel Quit $appId2
    Cawt Destroy
    exit 0
}
Cawt Destroy
