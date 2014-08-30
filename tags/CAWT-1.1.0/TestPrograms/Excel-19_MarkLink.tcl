# Test CawtExcel procedures related to marking and linking cells.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

# Open Excel, show the application window and create a workbook.
set appId [::Excel::Open true]
set workbookId [::Excel::AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-19_MarkLink"]
append xlsFile [::Excel::GetExtString $appId]
file delete -force $xlsFile

# Create two worksheets.
set refWorksheetId [::Excel::AddWorksheet $workbookId "ExcelRef"]
set worksheetId    [::Excel::AddWorksheet $workbookId "ExcelSource"]

# Generate a text file for testing the hyperlink capabilities.
set fileName [file join [pwd] "testOut" "Excel-19_MarkLink.txt"]
set fp [open $fileName "w"]
puts $fp "This is the linked text file."
close $fp

::Excel::SetCellValue $worksheetId    1 1 "Hyperlinked Value ExcelSource"
::Excel::SetCellValue $refWorksheetId 1 1 "Hyperlinked Value ExcelRef"
::Excel::SetCellValue $refWorksheetId 2 1 "Linked Value ExcelRef"
::Excel::SetColumnWidth $refWorksheetId 1 0

# Generic way to set hyperlinks.
::Excel::SetHyperlink $worksheetId 3 1 [format "file://%s" $fileName] "Hyperlink to file"
::Excel::SetHyperlink $worksheetId 4 1 [format "http://%s" "www.poSoft.de"] "Hyperlink to URL"

# Create a hyperlink to a file.
::Excel::SetHyperlinkToFile $worksheetId 5 1 $fileName "Hyperlink to file"

# Create a hyperlink to a cell in the same worksheet.
::Excel::SetHyperlinkToCell $worksheetId 1 1 $worksheetId 8 1
::Excel::SetHyperlinkToCell $worksheetId 1 1 $worksheetId 9 1 "Hyperlink to A1"

# Create a hyperlink to a cell in another worksheet.
::Excel::SetHyperlinkToCell $refWorksheetId 1 1 $worksheetId 12 1
::Excel::SetHyperlinkToCell $refWorksheetId 1 1 $worksheetId 13 1 "Hyperlink to ExcelRef!A1"

# Create an internal link to a cell in another worksheet.
::Excel::SetLinkToCell $refWorksheetId 2 1 $worksheetId 15 1

::Excel::SetColumnWidth $worksheetId 1 0

# Test adding comments.
::Excel::SetCellValue $worksheetId 1 3 "Cell with comment text"
set rangeId [::Excel::SelectCellByIndex $worksheetId 1 3 true]
::Excel::SetRangeComment $rangeId "This cell has a comment"
::Excel::SetRangeComment $rangeId "Overwritten comment text."

::Excel::SetCellValue $worksheetId 5 3 "Cell with comment image"
set rangeId [::Excel::SelectCellByIndex $worksheetId 5 3 true]
set commentId [::Excel::SetRangeComment $rangeId "Comment text." [file join [pwd] "testIn/wish.gif"]]
::Excel::SetCommentSize $commentId [::Cawt::CentiMetersToPoints 3] [::Cawt::CentiMetersToPoints 5]

::Excel::SetCommentDisplayMode $appId true true

# Test adding tooltips.
::Excel::SetCellValue $worksheetId 10 3 "Cell with tooltip"
set rangeId [::Excel::SelectCellByIndex $worksheetId 10 3 true]
::Excel::SetRangeTooltip $rangeId "Tooltip message" "Tooltip title"

::Excel::SetColumnWidth $worksheetId 3 0

puts "Saving as Excel file: $xlsFile"
::Excel::SaveAs $workbookId $xlsFile "" false

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
