# Test CawtWord procedures related to Word table management.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}
package require cawt

# Open new Word instance and show the application window.
set appId [Word OpenNew true]

# Delete Word file from previous test run.
file mkdir testOut
set wordFile [file join [pwd] "testOut" "Word-02_Table"]
append wordFile [Word GetExtString $appId]
file delete -force $wordFile

# Create a new document.
set docId [Word AddDocument $appId]


# Create a table with a header line.
set numRows 3
set numCols 5

Word AppendText $docId "A standard table with a header line:"
set table(1,Id) [Word AddTable [Word GetEndRange $docId] [expr $numRows+1] $numCols]
set table(1,Rows) [expr $numRows+1]

for { set c 1 } { $c <= $numCols } { incr c } {
    lappend headerList [format "Header-%d" $c]
}
Word SetHeaderRow $table(1,Id) $headerList

for { set r 1 } { $r <= $numRows } { incr r } {
    for { set c 1 } { $c <= $numCols } { incr c } {
        Word SetCellValue $table(1,Id) [expr $r+1] $c [format "R-%d C-%d" $r $c]
    }
}

# Create a table and change some properties.
set numRows 5
set numCols 2
Word AppendParagraph $docId
Word AppendText $docId "Another table with changed properties and added rows:"
set table(2,Id) [Word AddTable [Word GetEndRange $docId] $numRows $numCols 6]

for { set r 1 } { $r <= $numRows } { incr r } {
    for { set c 1 } { $c <= $numCols } { incr c } {
        Word SetCellValue $table(2,Id) $r $c [format "R-%d C-%d" $r $c]
    }
}

Word AddRow $table(2,Id) 1 2
Word AddRow $table(2,Id)
set table(2,Rows) [expr $numRows+3]

Word SetTableBorderLineStyle $table(2,Id)
Word SetTableBorderLineWidth $table(2,Id) wdLineWidth300pt

set rowRange [Word GetRowRange $table(2,Id) 1]
Word SetRangeFontBold $rowRange true
Word SetRangeBackgroundColor $rowRange 200 100 50

set colRange [Word GetColumnRange $table(2,Id) 2]
Word SetRangeFontItalic $colRange true

Word SetColumnWidth $table(2,Id) 1 [Cawt InchesToPoints 1]
Word SetColumnWidth $table(2,Id) 2 [Cawt CentiMetersToPoints 2.54]

# Read the number of rows and columns and check them.
set numRowsRead [Word GetNumRows $table(2,Id)]
set numColsRead [Word GetNumColumns $table(2,Id)]
Cawt CheckNumber [expr $numRows + 3] $numRowsRead "GetNumRows"
Cawt CheckNumber $numCols $numColsRead "GetNumColumns"

# Read back the contents of the table and insert them into a newly created table
# (which is 2 rows and 1 column larger than the original).
# Set all columns to an equal width and change the border style.
Word AppendParagraph $docId
Word AppendText $docId "Copy of table with changed borders:"
set table(3,Id) [Word AddTable [Word GetEndRange $docId] \
                [expr $numRows+2] [expr $numCols+1] 6]
set table(3,Rows) [expr $numRows+2]

set matrixList [Word GetMatrixValues $table(2,Id) 1 1 $numRows $numCols]
Word SetMatrixValues $table(3,Id) $matrixList 3 2

Word SetColumnsWidth $table(3,Id) 1 [expr $numCols+1] [Cawt InchesToPoints 1.9]
Word SetTableBorderLineStyle $table(3,Id) \
        wdLineStyleEmboss3D wdLineStyleDashDot

# Insert values into empty column starting at row 3.
set colList [list "Row-3" "Row-4" "Row-5" "Row-6"]
Word SetColumnValues $table(3,Id) 1 $colList 3

# Read back the values of the column starting at row 3.
set readList [Word GetColumnValues $table(3,Id) 1 3 [llength $colList]]
Cawt CheckList $colList $readList "GetColumnValues"

# Count the number of tables and return their identifiers.
set numTables [Word GetNumTables $docId]
Cawt CheckNumber 3 $numTables "GetNumTables"
for { set n 1 } { $n <= $numTables } {incr n } {
    set tableId [Word GetTableIdByIndex $docId $n]
    Cawt CheckNumber $table($n,Rows) [Word GetNumRows $tableId] "Table $n GetNumRows"
    Cawt Destroy $tableId
}

Word UpdateFields $docId

# Save document as Word file.
puts "Saving as Word file: $wordFile"
Word SaveAs $docId $wordFile

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Word Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
