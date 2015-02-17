# Test CawtWord procedures for handling text.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

# Open new Word instance and show the application window.
set appId [::Word::OpenNew true]

# Delete Word file from previous test run.
file mkdir testOut
set rootName [file join [pwd] "testOut" "Word-03_Text"]
set pdfFile  [format "%s%s" $rootName ".pdf"]
set wordFile [format "%s%s" $rootName [::Word::GetExtString $appId]]
file delete -force $pdfFile
file delete -force $wordFile

set msg1 "This is a italic line of text in italic.\n"
for { set i 0 } { $i < 20 } { incr i } {
    append msg3 "This is a large oops paragraph in bold. "
}

# Create a new document.
set docId [::Word::AddDocument $appId]

# Emtpy document has 1 paragraph character.
::Cawt::CheckNumber 1 [::Word::GetNumCharacters $docId] "Number of characters in empty document"

# Insert a short piece of text as one paragraph.
set range1 [::Word::AppendText $docId $msg1]
::Word::SetRangeFontItalic $range1 true
::Word::SetRangeFontSize $range1 12
::Word::SetRangeFontName $range1 "Courier"
::Word::SetRangeHighlightColorByEnum $range1 wdYellow

# 1 paragraph character + string
set expectedChars [expr 1 + [string length $msg1]]
::Cawt::CheckNumber $expectedChars [::Word::GetNumCharacters $docId] "Number of characters after adding text"

# Insert other short pieces of text with different underlinings.
set range2 [::Word::AppendText $docId "This is text with default underlining color.\n"]
::Word::SetRangeFontUnderline $range2

set range3 [::Word::AppendText $docId "This is text with orange underlining color.\n"]
::Word::SetRangeFontUnderline $range3 true wdColorLightOrange

# Insert a longer piece of text as one paragraph.
set range4 [::Word::AppendText $docId $msg3 true]
::Word::SetRangeFontBold $range4 true

# Test inserting different types of list.
set rangeId [::Word::AppendText $docId "Different types of lists" true]

set listRange [::Word::CreateRangeAfter $rangeId]
set listRange [::Word::InsertList $listRange [list "Unordered list entry 1" "Unordered list entry 2" "Unordered list entry 3"]]

set listRange [::Word::CreateRangeAfter $listRange]
set listRange [::Word::InsertList $listRange \
                   [list "Ordered list entry 1" "Ordered list entry 2" "Ordered list entry 3"] \
                   wdNumberGallery wdListListNumOnly]

# Insert lines of text. When we get to 7 inches from top of the
# document, insert a hard page break.
set pos [::Cawt::InchesToPoints 7]
while { true } {
    ::Word::AppendText $docId "More lines of text." true
    set endRange [::Word::GetEndRange $docId]
    if { $pos < [::Word::GetRangeInformation $endRange wdVerticalPositionRelativeToPage] } {
        break
    }
}

::Word::AddPageBreak $endRange

set rangeId [::Word::AppendText $docId "This is page 2." true]
::Word::AddParagraph $rangeId 10
::Word::AppendParagraph $docId 30
set rangeId [::Word::AppendText $docId "There must be two paragraphs before this line." true]


::Word::SetRangeStartIndex $rangeId "begin"
::Word::SetRangeEndIndex   $rangeId 5
::Word::SelectRange $rangeId
::Word::PrintRange $rangeId "Selected first 5 characters: "

# Save document as Word file.
puts "Saving as Word file: $wordFile"
::Word::SaveAs $docId $wordFile

puts "Saving as PDF file: $pdfFile"
# # Use in a catch statement, as PDF export is available only in Word 2007 an up.
set catchVal [ catch { ::Word::SaveAsPdf $docId $pdfFile } retVal]
if { $catchVal } {
    puts "Error: $retVal"
}

if { [lindex $argv 0] eq "auto" } {
    ::Word::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
