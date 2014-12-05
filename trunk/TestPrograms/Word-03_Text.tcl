# Test CawtWord procedures for handling text.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
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
::Word::SetRangeHighlightColorByEnum $range1 $::Word::wdYellow

# 1 paragraph character + string
set expectedChars [expr 1 + [string length $msg1]]
::Cawt::CheckNumber $expectedChars [::Word::GetNumCharacters $docId] "Number of characters after adding text"

# Insert other short pieces of text with different underlinings.
set range2 [::Word::AppendText $docId "This is text with default underlining color.\n"]
::Word::SetRangeFontUnderline $range2

set range3 [::Word::AppendText $docId "This is text with orange underlining color.\n"]
::Word::SetRangeFontUnderline $range3 true $::Word::wdColorLightOrange

# Insert a longer piece of text as one paragraph.
set range4 [::Word::AppendText $docId $msg3 true]
::Word::SetRangeFontBold $range4 true

# Generate a text file for testing the hyperlink capabilities.
set fileName [file join [pwd] "testOut" "Word-03_Text.txt"]
set fp [open $fileName "w"]
puts $fp "This is the text file linked from Word."
close $fp

::Word::AppendParagraph $docId
set rangeLink [::Word::AppendText $docId "Dummy"]
::Word::SetHyperlink $docId $rangeLink [format "file://%s" $fileName] "File Link"
::Word::AppendParagraph $docId

# Insert lines of text. When we get to 7 inches from top of the
# document, insert a hard page break.
set pos [::Cawt::InchesToPoints 7]
while { true } {
    ::Word::AppendText $docId "More lines of text." true
    set endRange [::Word::GetEndRange $docId]
    if { $pos < [::Word::GetRangeInformation $endRange $::Word::wdVerticalPositionRelativeToPage] } {
        break
    }
}

::Word::AddPageBreak $endRange

set rangeId [::Word::AppendText $docId "This is page 2." true]
::Word::AddParagraph $rangeId 10
::Word::AppendParagraph $docId 30
set rangeId [::Word::AppendText $docId "There must be two paragraphs before this line."]

::Word::SetRangeStartIndex $docId $rangeId "begin"
::Word::SetRangeEndIndex   $docId $rangeId 5
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
