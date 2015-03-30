# Test CawtExcel procedures to exchange data between Excel and the Windows clipboard.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

set outPath [file join [pwd] "testOut"]

set outFile [file join $outPath Excel-15_Clipboard]

# Create testOut directory, if it does not yet exist.
file mkdir testOut

# Open a new Excel instance, so we are able to get the extension string.
set appId [Excel OpenNew true]

set excelExt [Excel GetExtString $appId]

# Delete Excel output file from previous test run.
set xlsOutFile [format $outFile $excelExt]
file delete -force $xlsOutFile

# Create an Excel file with some test data.
set workbookId [Excel AddWorkbook $appId]
set headerList { "Col-1" "Col-2" "Col-3" "Col-4" }
set dataList {
    {"1" "2" "3" "None"}
    {"1.1" "1.2" "1.3" "Dot"}
    {"1,1" "1,2" "1,3" "Comma"}
    {"1|1" "1|2" "1|3" "Pipe"}
    {"1;1" "1;2" "1;3" "Semicolon"}
}

set worksheetId1 [Excel AddWorksheet $workbookId "ClipboardSource"]
Excel SetHeaderRow $worksheetId1 $headerList
Excel SetMatrixValues $worksheetId1 $dataList 2
set matrixList1 [Excel GetWorksheetAsMatrix $worksheetId1]

puts "Copy worksheet to clipboard"
Excel WorksheetToClipboard $worksheetId1 1 1  \
    [Excel GetLastUsedRow $worksheetId1] \
    [Excel GetLastUsedColumn $worksheetId1]

set worksheetId2 [Excel AddWorksheet $workbookId "ClipboardDest"]

puts "Copy clipboard to worksheet with offset"
Excel ClipboardToWorksheet $worksheetId2 3 2
set matrixList2 [Excel GetWorksheetAsMatrix $worksheetId2]

Cawt CheckMatrix $matrixList1 $matrixList2 "ClipboardToWorksheet"

Excel SaveAs $workbookId $xlsOutFile

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Excel Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
