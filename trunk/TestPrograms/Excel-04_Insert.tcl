# Test CawtExcel procedures for inserting data as rows, columns or matrices.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

# Number of test rows, columns and worksheets being generated.
set numRows  100
set numCols  10
set numWorksheets 3

# Generate list with test data.
for { set i 1 } { $i <= $numCols } { incr i } {
    lappend valList "$i"
}

# Test inserting data with the SetRowValues procedure.
proc InsertWithSetRowValues { appId workbookId rowList testName hideApp } {
    global numRows numCols numWorksheets

    set t1 [clock clicks -milliseconds]
    if { $hideApp } {
        ::Excel::Visible $appId false
    }
    set startRow 1
    set startCol 1
    for { set ws 0 } { $ws < $numWorksheets } { incr ws } {
        set worksheetId [::Excel::AddWorksheet $workbookId "$testName-$ws"]
        for { set row $startRow } { $row < [expr {$numRows+$startRow}] } { incr row } {
            ::Excel::SetRowValues $worksheetId $row $rowList $startCol
        }

        incr startRow 2
        incr startCol 1
    }
    if { $hideApp } {
        ::Excel::Visible $appId true
    }
    set t2 [clock clicks -milliseconds]
    puts "[expr $t2 - $t1] ms to set row values in $testName mode."
}

proc CheckWorksheets { workbookId } {
    global numRows numCols numWorksheets

    set startRow 1
    set startCol 1
    for { set ws $numWorksheets } { $ws >= 1 } { incr ws -1 } {
        set worksheetId [::Excel::GetWorksheetIdByIndex $workbookId $ws]
        set wsName [::Excel::GetWorksheetName $worksheetId]
        ::Cawt::CheckNumber $numRows [::Excel::GetNumUsedRows $worksheetId] \
                            "Number of used rows in $wsName"
        ::Cawt::CheckNumber $numCols [::Excel::GetNumUsedColumns $worksheetId] \
                            "Number of used columns in $wsName"
        ::Cawt::CheckNumber $startRow [::Excel::GetFirstUsedRow $worksheetId] \
                            "First used row in $wsName"
        ::Cawt::CheckNumber $startCol [::Excel::GetFirstUsedColumn $worksheetId] \
                            "First used column in $wsName"
        ::Cawt::CheckNumber [expr { $numRows + $startRow - 1 }] [::Excel::GetLastUsedRow $worksheetId] \
                            "Last used row in $wsName"
        ::Cawt::CheckNumber [expr { $numCols + $startCol - 1 }] [::Excel::GetLastUsedColumn $worksheetId] \
                            "Last used column in $wsName"

        incr startRow 2
        incr startCol 1
    }
}

# Test inserting data with the SetColumnValues procedure.
proc InsertWithSetColumnValues { appId workbookId rowList testName hideApp } {
    global numRows numWorksheets

    set t1 [clock clicks -milliseconds]
    if { $hideApp } {
        ::Excel::Visible $appId false
    }
    set startRow 1
    set startCol 1
    for { set ws 0 } { $ws < $numWorksheets } { incr ws } {
        set worksheetId [::Excel::AddWorksheet $workbookId "$testName-$ws"]
        for { set row $startRow } { $row < [expr {$numRows+$startRow}] } { incr row } {
            ::Excel::SetColumnValues $worksheetId $row $rowList $startCol
        }
        incr startRow 2
        incr startCol 1
    }
    if { $hideApp } {
        ::Excel::Visible $appId true
    }
    set t2 [clock clicks -milliseconds]
    puts "[expr $t2 - $t1] ms to set column values in $testName mode."
}

# Test inserting data with the SetMatrixValues procedure.
proc InsertWithSetMatrixValues { appId workbookId rowList testName hideApp } {
    global numRows numWorksheets

    set t1 [clock clicks -milliseconds]
    if { $hideApp } {
        ::Excel::Visible $appId false
    }
    for { set i 1 } { $i <= $numRows } { incr i } {
        lappend rangeList $rowList
    }
    set startRow 1
    set startCol 1
    for { set ws 0 } { $ws < $numWorksheets } { incr ws } {
        set worksheetId [::Excel::AddWorksheet $workbookId "$testName-$ws"]
        ::Excel::SetMatrixValues $worksheetId $rangeList $startRow $startCol
        incr startRow 2
        incr startCol 1
    }
    if { $hideApp } {
        ::Excel::Visible $appId true
    }
    set t2 [clock clicks -milliseconds]
    puts "[expr $t2 - $t1] ms to set row values in $testName mode."
}

# Open new instance of Excel and create a workbook.
set appId [::Excel::OpenNew]
set workbookId [::Excel::AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-04_Insert"]
append xlsFile [::Excel::GetExtString $appId]
file delete -force $xlsFile

# Perform test 1: Insert rows with Excel window visible.
InsertWithSetRowValues $appId $workbookId $valList "RowVisible" false
CheckWorksheets $workbookId

# Perform test 2: Insert rows with Excel window hidden.
InsertWithSetRowValues $appId $workbookId $valList "RowHidden" true

# Perform test 3: Insert matrix with Excel window visible.
InsertWithSetMatrixValues $appId $workbookId $valList "MatrixVisible" false

# Perform test 4: Insert matrix with Excel window hidden.
InsertWithSetMatrixValues $appId $workbookId $valList "MatrixHidden" true

# Perform test 5: Insert columns with Excel window visible.
InsertWithSetColumnValues $appId $workbookId $valList "ColumnVisible" false

# Perform test 6: Insert colums with Excel window hidden.
InsertWithSetColumnValues $appId $workbookId $valList "ColumnHidden" true

puts "Saving as Excel file: $xlsFile"
::Excel::SaveAs $workbookId $xlsFile

::Cawt::PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
