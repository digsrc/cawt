# Test CawtExcel procedure for diff'ing Excel files.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"

package require cawt

# Open Excel, so we can get the extension string.
set appId [::Excel::Open true]

set outPath [file join [pwd] "testOut"]
set xlsOutFile1 [file join $outPath Excel-17_Diff-Base[::Excel::GetExtString $appId]]
set xlsOutFile2 [file join $outPath Excel-17_Diff-New[::Excel::GetExtString $appId]]

# Create testOut directory, if it does not yet exist.
file mkdir testOut

# Delete Excel output file from previous test run.
file delete -force $xlsOutFile1
file delete -force $xlsOutFile2

# Create an Excel file with some test data.
set workbookId [::Excel::AddWorkbook $appId]
set headerList { "Col-1" "Col-2" "Col-3" "Col-4" }
set dataList {
    {"1" "2" "3" "None"}
    {"1.1" "1.2" "1.3" "Dot"}
    {"1,1" "1,2" "1,3" "Comma"}
    {"1|1" "1|2" "1|3" "Pipe"}
    {"1;1" "1;2" "1;3" "Semicolon"}
}

set worksheetId [::Excel::AddWorksheet $workbookId "DiffTest"]
::Excel::SetHeaderRow $worksheetId $headerList
::Excel::SetMatrixValues $worksheetId $dataList 2

# Test setting a formula.
set cell [::Excel::SelectCellByIndex $worksheetId 5 1 true]
$cell Formula "=TODAY()"

::Excel::SaveAs $workbookId $xlsOutFile1

# Change some values in the sheet and save to other Excel file.
::Excel::SetCellValue $worksheetId 1 1 12345
::Excel::SetCellValue $worksheetId 3 3 "ABCD"

::Excel::SaveAs $workbookId $xlsOutFile2

::Excel::Close $workbookId
::Excel::Quit $appId

set diffAppId [::Excel::DiffExcelFiles $xlsOutFile1 $xlsOutFile2 0 255 0]

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $diffAppId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
