# Test CawtExcel procedures to exchange data between Excel and Tablelist.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
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
set numRows 10
set numCols  5

# Column numbers of hidden columns.
# Note, that these are tablelist numbers starting at zero,
# while Excel column numbers start at 1.
set hiddenColumns { 0 3 }

# Generate header list with column names.
for { set c 1 } { $c <= $numCols } { incr c } {
    lappend headerList "Col-$c"
}

# Create frames for testing the transfer of full tablelists and for
# tablelists with hidden columns.
set fullFr .fullFr
set hideFr .hideFr
set frameList [list $fullFr $hideFr]
set modeList  [list "Full" "Hidden"]

ttk::labelframe $fullFr -padding 5 -text "Complete tablelists"
ttk::labelframe $hideFr -padding 5 -text "Tablelists with hidden columns"
pack $fullFr $hideFr -side left -fill both -expand true

set width 50

# Create 3 tablelist widgets for each mode.
foreach fr $frameList {
    ttk::labelframe $fr.frIn -text "Source table"
    pack $fr.frIn -side top -fill both -expand true
    tablelist::tablelist $fr.frIn.tl -width $width -height $numRows
    pack $fr.frIn.tl -side top -fill both -expand true

    ttk::labelframe $fr.frOut1 -text "Table with header"
    pack $fr.frOut1 -side top -fill both -expand true
    tablelist::tablelist $fr.frOut1.tl -width $width -height $numRows
    pack $fr.frOut1.tl -side top -fill both -expand true

    ttk::labelframe $fr.frOut2 -text "Table without header"
    pack $fr.frOut2 -side top -fill both -expand true
    tablelist::tablelist $fr.frOut2.tl -width $width -height $numRows
    pack $fr.frOut2.tl -side top -fill both -expand true
}

puts "Filling source tablelists with data"
foreach fr $frameList {
    Excel SetTablelistHeader $fr.frIn.tl $headerList
    Cawt CheckList $headerList [Excel GetTablelistHeader $fr.frIn.tl] "GetTablelistHeader"
    set matrixList [list]

    for { set row 1 } { $row <= $numRows } { incr row } {
        set rowList [list]
        for { set col 1 } { $col <= $numCols } { incr col } {
            lappend rowList [format "Cell_%d_%d" $row $col]
        }
        lappend matrixList $rowList
    }
    Excel SetTablelistValues $fr.frIn.tl $matrixList
    Cawt CheckMatrix $matrixList [Excel GetTablelistValues $fr.frIn.tl] "GetTablelistValues"
    update
}

foreach colNum $hiddenColumns {
    $hideFr.frIn.tl columnconfigure $colNum -hide true
}

# Open new instance of Excel and add a workbook.
set appId [Excel OpenNew]
set workbookId [Excel AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-08_Tablelist"]
append xlsFile [Excel GetExtString $appId]
file delete -force $xlsFile

foreach fr $frameList mode $modeList {
    # Transfer tablelist data with header information into Excel and vice versa.
    set useHeader true
    set worksheetId [Excel AddWorksheet $workbookId "${mode}_WithHeader"]

    set t1 [clock clicks -milliseconds]
    Excel TablelistToWorksheet $fr.frIn.tl $worksheetId $useHeader
    set t2 [clock clicks -milliseconds]
    puts "TablelistToWorksheet: [expr $t2 - $t1] ms (Mode $mode using header: $useHeader)."

    set t1 [clock clicks -milliseconds]
    Excel WorksheetToTablelist $worksheetId $fr.frOut1.tl $useHeader
    set t2 [clock clicks -milliseconds]
    puts "WorksheetToTablelist: [expr $t2 - $t1] ms (Mode $mode using header: $useHeader)."
    Cawt CheckMatrix $matrixList [Excel GetTablelistValues $fr.frOut1.tl] "GetTablelistValues"
    update

    # Transfer tablelist without header information into Excel and vice versa.
    set useHeader false
    set worksheetId [Excel AddWorksheet $workbookId "${mode}_NoHeader"]

    set t1 [clock clicks -milliseconds]
    Excel TablelistToWorksheet $fr.frIn.tl $worksheetId $useHeader
    set t2 [clock clicks -milliseconds]
    puts "TablelistToWorksheet: [expr $t2 - $t1] ms (Mode $mode using header: $useHeader)."

    set t1 [clock clicks -milliseconds]
    Excel WorksheetToTablelist $worksheetId $fr.frOut2.tl $useHeader
    set t2 [clock clicks -milliseconds]
    puts "WorksheetToTablelist: [expr $t2 - $t1] ms (Mode $mode using header: $useHeader)."
    Cawt CheckMatrix $matrixList [Excel GetTablelistValues $fr.frOut2.tl] "GetTablelistValues"
    update
}

puts "Saving as Excel file: $xlsFile"
Excel SaveAs $workbookId $xlsFile

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Excel Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
