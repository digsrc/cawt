# Test CawtExcel procedures to export Excel data to a HTML file.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}
package require cawt

# Open new instance of Excel and add a workbook.
set appId [Excel OpenNew]

set xlsFile [file join [pwd] "testIn" "SampleTable.xls"]

set outFileHtml1 [file join [pwd] "testOut" "Excel-22_Html1.html"]
set outFileHtml2 [file join [pwd] "testOut" "Excel-22_Html2.html"]

# Create testOut directory, if it does not yet exist.
file mkdir testOut

# Open the Excel file with a sample table.
set workbookId [Excel OpenWorkbook $appId $xlsFile true]
set worksheetId [Excel GetWorksheetIdByName $workbookId "SampleTable"]

puts "Copy worksheet to HTML file $outFileHtml1"
Excel WorksheetToHtmlFile $worksheetId $outFileHtml1

Excel Close $workbookId
Excel Quit $appId

puts "Copy Excel file to HTML file $outFileHtml2"
Excel ExcelFileToHtmlFile $xlsFile $outFileHtml2 "SampleTable" true

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Excel Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
