# Test CawtExcel procedures related to specifying number formats.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

# Open Excel, show the application window and create a workbook.
set appId [Excel Open true]
set workbookId [Excel AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-24_Format"]
append xlsFile [Excel GetExtString $appId]
file delete -force $xlsFile

set worksheetId [Excel GetWorksheetIdByIndex $workbookId 1]
Excel SetWorksheetName $worksheetId "NumberFormats"

Excel SetHeaderRow $worksheetId [list "Value as text" "Number format" "Display"]

# Insert a numeric value into a worksheet.
# Column 1 holds the value as a textual representation.
# Column 2 shows the number format specified.
# Column 3 is how Excel displays the value with the given number format.

proc InsertValue { worksheetId row value numberFormat } {
    Excel SetCellValue $worksheetId $row 1 $value
    Excel SetCellValue $worksheetId $row 2 $numberFormat
    Excel SetCellValue $worksheetId $row 3 $value "real" $numberFormat
}

puts "Using system separator   : [$appId UseSystemSeparators]"
puts "Using decimal separator  : [$appId DecimalSeparator]"
puts "Using thousands separator: [$appId ThousandsSeparator]"

set row 2

set numberFormat [Excel GetLangNumberFormat "0" "0000"]
InsertValue $worksheetId $row "1.0"     $numberFormat  ; incr row
InsertValue $worksheetId $row "0.54321" $numberFormat  ; incr row
InsertValue $worksheetId $row "12345.5" $numberFormat  ; incr row

set numberFormat [Excel GetLangNumberFormat "0" "0"]
InsertValue $worksheetId $row "1.0"     $numberFormat  ; incr row
InsertValue $worksheetId $row "0.54321" $numberFormat  ; incr row
InsertValue $worksheetId $row "12345.5" $numberFormat  ; incr row

Excel SetColumnsWidth $worksheetId 1 3

puts "Saving as Excel file: $xlsFile"
Excel SaveAs $workbookId $xlsFile "" false

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Excel Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
