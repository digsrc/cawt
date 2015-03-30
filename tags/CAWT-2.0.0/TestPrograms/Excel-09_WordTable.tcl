# Test CawtExcel procedures to exchange data between Excel and Word tables.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

# Number of header lines.
set numHeaders 1

# Number of test rows and columns being generated.
set numRows 10
set numCols  5

set totalRows [expr {$numRows + $numHeaders}]

# Generate header list with column names.
for { set c 1 } { $c <= $numCols } { incr c } {
    lappend headerList "Col-$c"
}

# Create 3 Word tables and fill the first with data.
set wordAppId [Word OpenNew true]
set docId [Word AddDocument $wordAppId]

Word AppendText $docId "Source table:" true
set tableIn [Word AddTable [Word GetEndRange $docId] $totalRows $numCols 1]
Word AppendParagraph $docId

Word AppendText $docId "Table with header:" true
set tableOut1 [Word AddTable [Word GetEndRange $docId] $totalRows $numCols 1]
Word AppendParagraph $docId

Word AppendText $docId "Table without header:" true
set tableOut2 [Word AddTable [Word GetEndRange $docId] $numRows $numCols 1]
Word AppendParagraph $docId

puts "Filling source table with data ..."
Word SetHeaderRow $tableIn $headerList

set matrixList [list]
for { set row 1 } { $row <= $numRows } { incr row } {
    set rowList [list]
    for { set col 1 } { $col <= $numCols } { incr col } {
        lappend rowList [format "Cell_%d_%d" $row $col]
    }
    Word SetRowValues $tableIn [expr {$row + $numHeaders}] $rowList
    lappend matrixList $rowList
}

# Open new instance of Excel and add a workbook.
set excelAppId [Excel OpenNew]
set workbookId [Excel AddWorkbook $excelAppId]

# Delete Excel and Word files from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-09_WordTable"]
append xlsFile [Excel GetExtString $excelAppId]
file delete -force $xlsFile

set docFile [file join [pwd] "testOut" "Excel-09_WordTable"]
append docFile [Word GetExtString $wordAppId]
file delete -force $docFile

# Transfer Word data with header information into Excel and vice versa.
set useHeader true

set worksheetId [Excel AddWorksheet $workbookId "WithHeader"]

set t1 [clock clicks -milliseconds]
Excel WordTableToWorksheet $tableIn $worksheetId $useHeader
set t2 [clock clicks -milliseconds]
puts "WordTableToWorksheet: [expr $t2 - $t1] ms (using header: $useHeader)."
Cawt CheckMatrix $matrixList \
    [Excel GetMatrixValues $worksheetId 2 1 [expr $numRows+1] $numCols] \
    "GetMatrixValues"

set t1 [clock clicks -milliseconds]
Excel WorksheetToWordTable $worksheetId $tableOut1 $useHeader
set t2 [clock clicks -milliseconds]
puts "WorksheetToWordTable: [expr $t2 - $t1] ms (using header: $useHeader)."


# Transfer Word data without header information into Excel and vice versa.
set useHeader false

set worksheetId [Excel AddWorksheet $workbookId "NoHeader"]

set t1 [clock clicks -milliseconds]
Excel WordTableToWorksheet $tableIn $worksheetId $useHeader
set t2 [clock clicks -milliseconds]
puts "WordTableToWorksheet: [expr $t2 - $t1] ms (using header: $useHeader)."
Cawt CheckMatrix $matrixList \
    [Excel GetMatrixValues $worksheetId 1 1 $numRows $numCols] \
    "GetMatrixValues"

set t1 [clock clicks -milliseconds]
Excel WorksheetToWordTable $worksheetId $tableOut2 $useHeader
set t2 [clock clicks -milliseconds]
puts "WorksheetToWordTable: [expr $t2 - $t1] ms (using header: $useHeader)."

puts "Saving as Excel file: $xlsFile"
Excel SaveAs $workbookId $xlsFile

puts "Saving as Word file : $docFile"
Word SaveAs $docId $docFile

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Excel Quit $excelAppId
    Word  Quit $wordAppId
    Cawt Destroy
    exit 0
}
Cawt Destroy
