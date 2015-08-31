# Test CawtExcel procedures for handling sparse matrices.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}
package require cawt

# Generate list with sparse data.
set testMatrix [list \
    [list  1  2  3 "" ""  6  7  8  9 10] \
    [list "" ""  3  4  5  6  7 "" "" 10] \
    [list  1  2  "" 4  5 "" "" ""  9 ""] \
    [list  1 ""  3 ""  5 ""  7 ""  9 ""] \
]

set numRows [llength $testMatrix]
set numCols 10 

set numWorksheets 3

proc CheckWorksheets { workbookId testName } {
    global numRows numCols numWorksheets

    set startRow 1
    set startCol 1
    for { set ws 0 } { $ws < $numWorksheets } { incr ws } {
        set wsName "$testName-$ws"
        set worksheetId [Excel GetWorksheetIdByName $workbookId $wsName]
        Cawt CheckNumber $numRows [Excel GetNumUsedRows $worksheetId] \
                         "Number of used rows in $wsName"
        Cawt CheckNumber $numCols [Excel GetNumUsedColumns $worksheetId] \
                         "Number of used columns in $wsName"
        Cawt CheckNumber $startRow [Excel GetFirstUsedRow $worksheetId] \
                         "First used row in $wsName"
        Cawt CheckNumber $startCol [Excel GetFirstUsedColumn $worksheetId] \
                         "First used column in $wsName"
        Cawt CheckNumber [expr { $numRows + $startRow - 1 }] [Excel GetLastUsedRow $worksheetId] \
                         "Last used row in $wsName"
        Cawt CheckNumber [expr { $numCols + $startCol - 1 }] [Excel GetLastUsedColumn $worksheetId] \
                         "Last used column in $wsName"

        incr startRow 2
        incr startCol 1
    }
}

# Test inserting data with the SetRowValues procedure.
proc InsertWithSetRowValues { workbookId matrixList testName } {
    global numWorksheets

    set startRow 1
    set startCol 1
    for { set ws 0 } { $ws < $numWorksheets } { incr ws } {
        set worksheetId [Excel AddWorksheet $workbookId "$testName-$ws"]
        set row $startRow
        foreach rowList $matrixList {
            Excel SetRowValues $worksheetId $row $rowList $startCol
            incr row
        }
        incr startRow 2
        incr startCol 1
    }
}

# Test inserting data with the SetMatrixValues procedure.
proc InsertWithSetMatrixValues { workbookId matrixList testName } {
    global numWorksheets

    set startRow 1
    set startCol 1
    for { set ws 0 } { $ws < $numWorksheets } { incr ws } {
        set worksheetId [Excel AddWorksheet $workbookId "$testName-$ws"]
        Excel SetMatrixValues $worksheetId $matrixList $startRow $startCol
        incr startRow 2
        incr startCol 1
    }
}

# Open new instance of Excel and create a workbook.
set appId [Excel OpenNew]
set workbookId [Excel AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-18_SparseMatrix"]
append xlsFile [Excel GetExtString $appId]
file delete -force $xlsFile

# Perform test 1: Insert rows with procedure SetRowValues.
InsertWithSetRowValues $workbookId $testMatrix "SetRowValues"
CheckWorksheets $workbookId "SetRowValues"

# Perform test 2: Insert columns with procedure SetColumnValues.
InsertWithSetMatrixValues $workbookId $testMatrix "SetMatrixValues"
CheckWorksheets $workbookId "SetMatrixValues"

puts "Saving as Excel file: $xlsFile"
Excel SaveAs $workbookId $xlsFile

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Excel Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
