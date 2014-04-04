# Test CawtWord procedures related to Word table management.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

# Open new Word instance and show the application window.
set appId [::Word::OpenNew true]

# Delete Word file from previous test run.
file mkdir testOut
set wordFile [file join [pwd] "testOut" "Word-02_Table"]
append wordFile [::Word::GetExtString $appId]
file delete -force $wordFile

# Create a new document.
set docId [::Word::AddDocument $appId]

# Create a table with a header line.
set numRows 3
set numCols 5

set endRange [::Word::AppendParagraph $docId "A standard table with a header line:"]
set table1Id [::Word::AddTable $docId $endRange [expr $numRows+1] $numCols 1]

for { set c 1 } { $c <= $numCols } { incr c } {
    lappend headerList [format "Header-%d" $c]
}
::Word::SetHeaderRow $table1Id $headerList

for { set r 1 } { $r <= $numRows } { incr r } {
    for { set c 1 } { $c <= $numCols } { incr c } {
        ::Word::SetCellValue $table1Id [expr $r+1] $c [format "R-%d C-%d" $r $c]
    }
}

# Create a table and change some properties.
set numRows 5
set numCols 2
set endRange [::Word::AppendParagraph $docId "Another table with changed properties:"]
set table2Id [::Word::AddTable $docId $endRange $numRows $numCols 6]

for { set r 1 } { $r <= $numRows } { incr r } {
    for { set c 1 } { $c <= $numCols } { incr c } {
        ::Word::SetCellValue $table2Id $r $c [format "R-%d C-%d" $r $c]
    }
}

::Word::AddRow $table2Id 1
::Word::AddRow $table2Id

::Word::SetTableBorderLineStyle $table2Id
::Word::SetTableBorderLineWidth $table2Id $::Word::wdLineWidth300pt

set rowRange [::Word::GetRowRange $table2Id 1]
::Word::SetRangeFontBold $rowRange true
::Word::SetRangeBackgroundColor $rowRange 200 100 50

set colRange [::Word::GetColumnRange $table2Id 2]
::Word::SetRangeFontItalic $colRange true

::Word::SetColumnWidth $table2Id 1 [::Cawt::InchesToPoints 1]
::Word::SetColumnWidth $table2Id 2 [::Cawt::CentiMetersToPoints 2.54]

# Read the number of rows and columns and check them.
set numRowsRead [::Word::GetNumRows $table2Id]
set numColsRead [::Word::GetNumColumns $table2Id]
if { [expr $numRows + 2] != $numRowsRead } {
    puts "Error: Number of rows not identical ($numRows vs. $numRowsRead)"
}
if { $numCols != $numColsRead } {
    puts "Error: Number of columns not identical ($numCols vs. $numColsRead)"
}

# Read back the contents of the table and insert them into a newly created table
# (which is 2 rows and 1 column larger than the original).
# Set all columns to an equal width and change the border style.
set endRange [::Word::AppendParagraph $docId "Copy of table with changed borders:"]
set table3Id [::Word::AddTable $docId $endRange \
              [expr $numRows+2] [expr $numCols+1] 6]

set matrixList [::Word::GetMatrixValues $table2Id 1 1 $numRows $numCols]
::Word::SetMatrixValues $table3Id $matrixList 3 2

::Word::SetColumnsWidth $table3Id 1 [expr $numCols+1] [::Cawt::InchesToPoints 1.9]
::Word::SetTableBorderLineStyle $table3Id \
        $::Word::wdLineStyleEmboss3D $::Word::wdLineStyleDashDot

# Insert values into empty column starting at row 3.
set colList [list "Row-3" "Row-4" "Row-5" "Row-6"]
::Word::SetColumnValues $table3Id 1 $colList 3

# Read back the values of the column starting at row 3.
set readList [::Word::GetColumnValues $table3Id 1 3 [llength $colList]]
foreach r $readList c $colList {
    if { $r ne $c } {
        puts "Error: Read-back list is not identical to original list"
        puts $readList
    }
}

# Count the number of tables and return their identifiers.
set numTables [::Word::GetNumTables $docId]
::Cawt::CheckNumber 3 $numTables "GetNumTables"
for { set n 1 } { $n <= $numTables } {incr n } {
    set tableId [::Word::GetTableIdByIndex $docId $n]
    puts "  Table $n has [::Word::GetNumRows $tableId] rows"
}

# Test inserting an image.
set endRange [::Word::AppendParagraph $docId \
             "The wish lamp cropped at the right side:"]
set imgId [::Word::InsertImage $endRange [file join [pwd] "testIn/wish.gif"]]

# CropImage imgId cropBottom cropTop cropLeft cropRight
::Word::CropImage $imgId 0 0 0 [::Cawt::CentiMetersToPoints 0.3]

::Word::UpdateFields $docId

# Save document as Word file.
puts "Saving as Word file: $wordFile"
::Word::SaveAs $docId $wordFile

if { [lindex $argv 0] eq "auto" } {
    ::Word::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
