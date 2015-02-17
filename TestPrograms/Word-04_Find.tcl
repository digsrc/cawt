# Test CawtWord procedures related to search and replace functionality.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

# Open new Word instance and show the application window.
set appId [::Word::Open true]

# Delete Word file from previous test run.
file mkdir testOut
set wordFile [file join [pwd] "testOut" "Word-04_Find"]
append wordFile [::Word::GetExtString $appId]
file delete -force $wordFile

set inFile [file join [pwd] "testOut" "Word-03_Text"]
append inFile [::Word::GetExtString $appId]

# Open an existing document. Set compatibility mode to Word 2003.
set inDocId  [::Word::OpenDocument $appId $inFile]
::Word::SetCompatibilityMode $inDocId wdWord2003

set range [::Word::GetStartRange $inDocId]
if { [::Word::GetRangeStartIndex $range] != 0 || \
     [::Word::GetRangeEndIndex   $range] != 0 } {
    puts "Error: Start range not correct"
    ::Word::PrintRange $range
}
if { ! [::Word::FindString $range "italic"] } {
    puts "Error: Word \"italic\" not listed in Word-Document"
}

set endIndex [::Word::GetRangeEndIndex $range]
set range [::Word::ExtendRange $range 0 500]
::Cawt::CheckNumber [expr $endIndex + 500] [::Word::GetRangeEndIndex $range] "End index of extended range"
::Word::PrintRange $range "Extended range:"
::Word::ReplaceString $range "italic" "yellow" "one"

set range [::Word::ExtendRange $range 0 end]
::Word::PrintRange $range "Extended range:"
::Word::ReplaceString $range "oops " "" "all"

set range [::Word::ExtendRange $range 0 end]
::Word::PrintRange $range "Extended range:"
::Word::ReplaceString $range "lines" "rows" "all"

::Word::ReplaceByProc [::Word::GetStartRange $inDocId] "paragraph" \
                        ::Word::SetRangeFontItalic true
# TODO This does not work
#::Word::ReplaceByProc [::Word::GetStartRange $inDocId] "paragraph" \
#                        ::Word::SetRangeHighlightColorByEnum wdYellow

::Word::InsertText $inDocId "Inserted text at beginning of document\n"

# Save document as Word file.
puts "Saving as Word file: $wordFile"
::Word::SaveAs $inDocId $wordFile

# Get number of open documents.
set numDocs [::Word::GetNumDocuments $appId]
::Cawt::CheckNumber 1 $numDocs "Number of open documents"

set newDocId [::Word::OpenDocument $appId $inFile]
set numDocs [::Word::GetNumDocuments $appId]
::Cawt::CheckNumber 2 $numDocs "Number of open documents"
for { set i 1 } { $i <= $numDocs } { incr i } {
    set docId [::Word::GetDocumentIdByIndex $appId $i]
    puts "File-$i: [::Word::GetDocumentName $docId]"
}
::Word::Close $newDocId

if { [lindex $argv 0] eq "auto" } {
    ::Word::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
