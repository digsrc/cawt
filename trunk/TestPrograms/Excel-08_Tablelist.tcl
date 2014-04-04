# Test CawtExcel procedures to exchange data between Excel and Tablelist.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"

# We need to explicitly load the tablelist package.
package require Tk

set retVal [catch {package require tablelist} version]
if { $retVal != 0 } {
    puts "Test not performed. Tablelist is not available."
    exit 0
}

package require cawt

# Number of test rows and columns being generated.
set numRows  10
set numCols   5

# Generate header list with column names.
for { set c 1 } { $c <= $numCols } { incr c } {
    lappend headerList "Col-$c"
}

# Create 3 tablelist widgets and fill the first with data.
ttk::labelframe .frIn -text "Source table:"
pack .frIn -side top -fill both -expand 1
tablelist::tablelist .frIn.tl -width 100
pack .frIn.tl -side top -fill both -expand 1

ttk::labelframe .frOut1 -text "Table with header:"
pack .frOut1 -side top -fill both -expand 1
tablelist::tablelist .frOut1.tl -width 100
pack .frOut1.tl -side top -fill both -expand 1

ttk::labelframe .frOut2 -text "Table without header:"
pack .frOut2 -side top -fill both -expand 1
tablelist::tablelist .frOut2.tl -width 100
pack .frOut2.tl -side top -fill both -expand 1

puts "Filling source tablelist with data"
::Excel::SetTablelistHeader .frIn.tl $headerList
::Cawt::CheckList $headerList [::Excel::GetTablelistHeader .frIn.tl] "GetTablelistHeader"
set matrixList [list]

for { set row 1 } { $row <= $numRows } { incr row } {
    set rowList [list]
    for { set col 1 } { $col <= $numCols } { incr col } {
        lappend rowList [format "Cell_%d_%d" $row $col]
    }
    lappend matrixList $rowList
}
::Excel::SetTablelistValues .frIn.tl $matrixList
::Cawt::CheckMatrix $matrixList [::Excel::GetTablelistValues .frIn.tl] "GetTablelistValues"
update

# Open new instance of Excel and add a workbook.
set appId [::Excel::OpenNew]
set workbookId [::Excel::AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-08_Tablelist"]
append xlsFile [::Excel::GetExtString $appId]
file delete -force $xlsFile

# Transfer tablelist data with header information into Excel and vice versa.
set useHeader true

set worksheetId [::Excel::AddWorksheet $workbookId "WithHeader"]

set t1 [clock clicks -milliseconds]
::Excel::TablelistToWorksheet .frIn.tl $worksheetId $useHeader
set t2 [clock clicks -milliseconds]
puts "TablelistToWorksheet: [expr $t2 - $t1] ms (using header: $useHeader)."

set t1 [clock clicks -milliseconds]
::Excel::WorksheetToTablelist $worksheetId .frOut1.tl $useHeader
set t2 [clock clicks -milliseconds]
puts "WorksheetToTablelist: [expr $t2 - $t1] ms (using header: $useHeader)."
update


# Transfer tablelist without header information into Excel and vice versa.
set useHeader false

set worksheetId [::Excel::AddWorksheet $workbookId "NoHeader"]

set t1 [clock clicks -milliseconds]
::Excel::TablelistToWorksheet .frIn.tl $worksheetId $useHeader
set t2 [clock clicks -milliseconds]
puts "TablelistToWorksheet: [expr $t2 - $t1] ms (using header: $useHeader)."

set t1 [clock clicks -milliseconds]
::Excel::WorksheetToTablelist $worksheetId .frOut2.tl $useHeader
set t2 [clock clicks -milliseconds]
puts "WorksheetToTablelist: [expr $t2 - $t1] ms (using header: $useHeader)."
update

puts "Saving as Excel file: $xlsFile"
::Excel::SaveAs $workbookId $xlsFile

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
