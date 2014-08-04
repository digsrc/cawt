# Test CawtWord procedures needed for generating a test report.
# It is assumed that the test data is read from an external data sink (ex. file),
# there are images to be inserted that are in a format Word does not know about,
# and a summary should be printed at the first page.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt
package require Tk

# Open new Word instance and show the application window.
set appId [::Word::OpenNew true]

# Delete Word file from previous test run.
file mkdir testOut
set rootName [file join [pwd] "testOut" "Word-05_Report"]
set wordFile [format "%s%s" $rootName [::Word::GetExtString $appId]]
file delete -force $wordFile

# Create a new document.
set docId [::Word::AddDocument $appId]

# Switch off spell and grammatic checking.
::Word::ToggleSpellCheck $appId false

set numTestCases   3
set testCaseTmpl   "Test case %d"
set testResultTmpl "Test case result: %s"
set testImg        [file join [pwd] "testIn/wish.gif"]
set numTestsOk     0
set numTestsFail   0
set numTests       0

# Simulate inserting the test data.
for { set t 1 } { $t <= $numTestCases } { incr t } {
    # Add chapter title.
    set title [format $testCaseTmpl $t]
    set titleRange [::Word::AppendText $docId $title]
    ::Word::SetRangeStyle $titleRange $::Word::wdStyleHeading1
    set titleRange [::Word::AppendParagraph $docId]

    # Add test result.
    if { [expr $t % 3] != 0 } {
        set success "OK"
        incr numTestsOk
    } else {
        set success "FAIL"
        incr numTestsFail
    }
    set result [format $testResultTmpl $success]
    set resultRange [::Word::AppendText $docId $result]
    ::Word::SetRangeStyle $titleRange $::Word::wdStyleBodyText
    set resultRange [::Word::AppendParagraph $docId]

    # Add the image related to current test case via the clipboard.
    set phImg [image create photo -file $testImg]
    ::Cawt::ImgToClipboard $phImg
    after 200
    $resultRange Paste
    image delete $phImg

    # Add a page break for new test case.
    ::Word::AddPageBreak $resultRange

    incr numTests
}

# Print summary of test suite at the beginning of the document.
set startRange [::Word::InsertText $docId "Summary of performed tests\n" $::Word::wdStyleTitle]

append summary "Number of test cases      : $numTests\n"
append summary "Number of successful tests: $numTestsOk\n"
append summary "Number of failed tests    : $numTestsFail\n"
set sumRange [::Word::AddText $docId $startRange $summary $::Word::wdStylePlainText]

::Word::SelectRange $sumRange
set checkRange [::Word::GetSelectionRange $docId]
::Cawt::CheckNumber [::Word::GetRangeStartIndex $sumRange] [::Word::GetRangeStartIndex $checkRange] "Start index of selected range"
::Cawt::CheckNumber [::Word::GetRangeEndIndex $sumRange] [::Word::GetRangeEndIndex $checkRange] "End index of selected range"

::Word::AddPageBreak $sumRange

# Save document as Word file.
puts "Saving as Word file: $wordFile"
::Word::SaveAs $docId $wordFile

if { [lindex $argv 0] eq "auto" } {
    ::Word::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
