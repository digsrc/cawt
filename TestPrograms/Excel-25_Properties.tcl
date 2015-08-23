# Test CawtExcel procedures related to property handling.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

# Open Excel, show the application window and create a workbook.
set appId [Excel OpenNew true]
set workbookId [Excel AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-25_Properties"]
append xlsFile [Excel GetExtString $appId]
file delete -force $xlsFile

# Set some builtin and custom properties and check their values.
Cawt SetDocumentProperty $workbookId "Author"     "Paul Obermeier"
Cawt SetDocumentProperty $workbookId "Company"    "poSoft"
Cawt SetDocumentProperty $workbookId "Title"      $xlsFile
Cawt SetDocumentProperty $workbookId "CustomProp" "CustomValue"

Cawt CheckString "Paul Obermeier" [Cawt GetDocumentProperty $workbookId "Author"]     "Property Author"
Cawt CheckString "poSoft"         [Cawt GetDocumentProperty $workbookId "Company"]    "Property Company"
Cawt CheckString $xlsFile         [Cawt GetDocumentProperty $workbookId "Title"]      "Property Title"
Cawt CheckString "CustomValue"    [Cawt GetDocumentProperty $workbookId "CustomProp"] "Property CustomProp"

set worksheetId [Excel GetWorksheetIdByIndex $workbookId 1]

Excel SetHeaderRow $worksheetId "BuiltinProperties" 1 1
Excel SetHeaderRow $worksheetId "CustomProperties" 1 3
Excel SetHeaderRow $worksheetId [list "Name" "Value" "Name" "Value"] 2 1
set rangeId [Excel SelectRangeByIndex $worksheetId 1 1 1 2]
Excel SetRangeMergeCells $rangeId true
set rangeId [Excel SelectRangeByIndex $worksheetId 1 3 1 4]
Excel SetRangeMergeCells $rangeId true

# Get all builtin and custom properties and insert them into the worksheet.
set row 3
set col 1
foreach propertyName [Cawt GetDocumentProperties $workbookId "Builtin"] {
    Excel SetCellValue $worksheetId $row [expr $col + 0] $propertyName
    Excel SetCellValue $worksheetId $row [expr $col + 1] [Cawt GetDocumentProperty $workbookId $propertyName]
    incr row
}

set row 3
set col 3
foreach propertyName [Cawt GetDocumentProperties $workbookId "Custom"] {
    Excel SetCellValue $worksheetId $row [expr $col + 0] $propertyName
    Excel SetCellValue $worksheetId $row [expr $col + 1] [Cawt GetDocumentProperty $workbookId $propertyName]
    incr row
}

Excel SetColumnsWidth $worksheetId 1 4

puts "Saving as Excel file: $xlsFile"
Excel SaveAs $workbookId $xlsFile "" false

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Excel Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
