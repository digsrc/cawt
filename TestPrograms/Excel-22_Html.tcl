# Test CawtExcel procedures to exchange data between Excel and Tablelist.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt


# Open new instance of Excel and add a workbook.
set appId [Excel OpenNew]

set outFileHtml1 [file join [pwd] "testOut" "Excel-22_Html1.html"]
set outFileHtml2 [file join [pwd] "testOut" "Excel-22_Html2.html"]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-22_Html"]
append xlsFile [Excel GetExtString $appId]
file delete -force $xlsFile

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

set worksheetId [Excel AddWorksheet $workbookId "HtmlTest"]
Excel SetHeaderRow $worksheetId $headerList
Excel SetMatrixValues $worksheetId $dataList 2

Excel SetRangeFillColor [Excel GetCellIdByIndex $worksheetId 1 1] 255   0 0
Excel SetRangeFillColor [Excel SelectRangeByIndex $worksheetId 5 1 5 4]   0 255 0

puts "Copy worksheet to HTML file $outFileHtml1"
Excel WorksheetToHtmlFile $worksheetId $outFileHtml1 true

Excel SaveAs $workbookId $xlsFile
Excel Close $workbookId
Excel Quit $appId

puts "Copy Excel file to HTML file $outFileHtml2"
Excel ExcelFileToHtmlFile $xlsFile $outFileHtml2 "HtmlTest" true true

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Excel Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
