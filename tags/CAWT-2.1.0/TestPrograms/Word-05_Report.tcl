# Test CawtWord procedures needed for generating a test report.
# It is assumed that the test data is read from an external data sink (ex. file),
# there are images to be inserted that are in a format Word does not know about,
# and a summary should be printed at the first page.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}
package require cawt
package require Tk

# Open new Word instance and show the application window.
set appId [Word OpenNew true]

# Delete Word file from previous test run.
file mkdir testOut
set rootName [file join [pwd] "testOut" "Word-05_Report"]
set wordFile [format "%s%s" $rootName [Word GetExtString $appId]]
file delete -force $wordFile

set inFile [file join [pwd] "testIn" "ReportTemplate.doc"]

# Open the report template document.
if { 1 } {
    set docId [Word OpenDocument $appId $inFile]
} else {
    set docId [Word AddDocument $appId]
}

# Switch off spell and grammatic checking.
Word ToggleSpellCheck $appId false

# Configure the Figure captions inserted later.
Word ConfigureCaption $appId $Word::wdCaptionFigure 1 true

set numTestSuites  3
set numTestCases   2
set testSuiteTmpl  "Test suite %d"
set testCaseTmpl   "Test case %d"
set testResultTmpl "Test case result: %s"
set testImg        [file join [pwd] "testIn/wish.gif"]
set numTestsOk     0
set numTestsFail   0
set numTests       0

# Simulate inserting the test data at the end of the document.
Word AppendParagraph $docId
for { set s 1 } { $s <= $numTestSuites } { incr s } {
    # Add test suite name as heading type 1.
    set title [format $testSuiteTmpl $s]
    set titleRange [Word AppendText $docId $title true]
    Word SetRangeStyle $titleRange wdStyleHeading1
    #Word AppendParagraph $docId
    Word AppendText $docId "Test cases of test suite $s\n"
    for { set t 1 } { $t <= $numTestCases } { incr t } {
        # Add test case name as heading type 2.
        set title [format $testCaseTmpl $t]
        set titleRange [Word AppendText $docId $title]
        Word SetRangeStyle $titleRange wdStyleHeading2
        set titleRange [Word AppendParagraph $docId]

        # Add test result.
        if { [expr $t % 2] != 0 } {
            set success "OK"
            incr numTestsOk
        } else {
            set success "FAIL"
            incr numTestsFail
        }
        set result [format $testResultTmpl $success]
        set resultRange [Word AppendText $docId $result true wdStyleBodyText]
        set bookmarkIds($s.$t) [Word AddBookmark $resultRange "Bookmark-$s-$t"]

        # Add the image related to current test case via the clipboard.
        set phImg [image create photo -file $testImg]
        Cawt ImgToClipboard $phImg
        after 200
        [Word GetEndRange $docId] Paste
        image delete $phImg
        set resultRange [Word AppendParagraph $docId]

        # Add a caption under the image.
        Word InsertCaption $resultRange $Word::wdCaptionFigure ": Test image $s.$t"
        set resultRange [Word AppendParagraph $docId]

        incr numTests
    }
}

# Replace keyword %TITLE% with actual title string.
Word ReplaceString [Word GetStartRange $docId] "%TITLE%" "Test report"

# Replace keyword %SUMMARY% with summary of test suite.
set summaryRange [Word GetStartRange $docId]
Word ReplaceString $summaryRange "%SUMMARY%" "Summary of performed tests"
set summaryRange [Word AddParagraph $summaryRange]
Word SetRangeStyle $summaryRange wdStyleHeading1

set summary "\n"
append summary "Number of test suites     : $numTestSuites\n"
append summary "Number of test cases      : $numTests\n"
append summary "Number of successful tests: $numTestsOk\n"
append summary "Number of failed tests    : $numTestsFail\n"
set sumRange [Word AddText $summaryRange $summary false wdStylePlainText]
set sumRange [Word AddParagraph $sumRange]

Word SelectRange $sumRange
set checkRange [Word GetSelectionRange $docId]
Cawt CheckNumber [Word GetRangeStartIndex $sumRange] [Word GetRangeStartIndex $checkRange] "Start index of selected range"
Cawt CheckNumber [Word GetRangeEndIndex $sumRange] [Word GetRangeEndIndex $checkRange] "End index of selected range"

set tableRange [Word CreateRangeAfter $sumRange]
set tableId [Word AddTable $tableRange [expr [array size bookmarkIds] + 1] 2]
Word SetTableBorderLineStyle $tableId
Word SetHeaderRow $tableId [list "Requirement" "Test Case"]
set req 1
set row 2
foreach key [lsort [array names bookmarkIds]] {
    Word SetCellValue $tableId $row 1 "Requirement $req"
    Word SetCellValue $tableId $row 2 [Word GetBookmarkName $bookmarkIds($key)]
    set cellRange [Word GetCellRange $tableId $row 2]
    Word SetLinkToBookmark $cellRange $bookmarkIds($key) "Link to test case $key"
    Cawt Destroy $cellRange
    incr row
    incr req
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
