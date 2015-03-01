# Test CawtExcel procedures for setting and getting cell values.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

set startCol 1

# Open Excel, show the application window and create a workbook.
set appId [Excel Open true]
set workbookId [Excel AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-16_SetGet"]
append xlsFile [Excel GetExtString $appId]
file delete -force $xlsFile

# Select the first - already existing - worksheet,
# set its name and fill it with data.
set worksheetId [Excel GetWorksheetIdByIndex $workbookId 1]
Excel SetWorksheetName $worksheetId "SetGet"

# Insert values with the different SetCellValue procedures.
Excel SetCellValue $worksheetId 1 1  "SetCellValue text"
Excel SetCellValue $worksheetId 2 1  1
Excel SetCellValue $worksheetId 3 1  1.4
Excel SetCellValue $worksheetId 4 1  1.6

Excel SetCellValue $worksheetId 1 2  "SetCellValue int"
Excel SetCellValue $worksheetId 2 2  1   "int"
Excel SetCellValue $worksheetId 3 2  1.4 "int"
Excel SetCellValue $worksheetId 4 2  1.6 "int"

Excel SetCellValue $worksheetId 1 3  "SetCellValue real"
Excel SetCellValue $worksheetId 2 3  1   "real"
Excel SetCellValue $worksheetId 3 3  1.4 "real"
Excel SetCellValue $worksheetId 4 3  1.6 "real"

Excel FormatHeaderRow  $worksheetId 1  1 3
Excel SetColumnsWidth  $worksheetId 1 3  0

Cawt CheckString "1.6" [Excel GetCellValue $worksheetId 4 1 "text"] "GetCellValue text"
Cawt CheckNumber 1     [Excel GetCellValue $worksheetId 4 2 "int"]  "GetCellValue int"
Cawt CheckNumber 1.6   [Excel GetCellValue $worksheetId 4 3 "real"] "GetCellValue real"

puts "Saving as Excel file: $xlsFile"
Excel SaveAs $workbookId $xlsFile

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Excel Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
