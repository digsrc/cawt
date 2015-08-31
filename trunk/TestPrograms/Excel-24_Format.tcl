# Test CawtExcel procedures related to specifying number formats.
# If called without options, the system separators are used.
# If "--English" is specified, a dot is used as floating point separator and a comma as thousands separator.
# If "--German"  is specified, a comma is used as floating point separator and a dot as thousands separator.
#
# Note, that if using one of the 2 options, the standard settings regarding separators are changed globally
# for the installed Excel application.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}
package require cawt

# Insert and read back a numeric value into a worksheet.
# Column 1 holds the value as a textual representation.
# Column 2 shows the number format specified.
# Column 3 is how Excel displays the value with the given number format.

proc InsertValue { worksheetId row value numberFormat } {
    Excel SetCellValue $worksheetId $row 1 $value
    Excel SetCellValue $worksheetId $row 2 $numberFormat
    Excel SetCellValue $worksheetId $row 3 $value "real" $numberFormat

    Cawt CheckString $value        [Excel GetCellValue $worksheetId $row 1]        "Value as text"
    Cawt CheckString $numberFormat [Excel GetCellValue $worksheetId $row 2]        "Number format"
    Cawt CheckNumber $value        [Excel GetCellValue $worksheetId $row 3 "real"] "Display"
}

set appId [Excel Open true]
set workbookId [Excel AddWorkbook $appId]

if { [lsearch -nocase $argv "--German"] >= 0 } {
    $appId UseSystemSeparators false
    set floatSep     ","
    set thousandsSep "."
    $appId DecimalSeparator   $floatSep
    $appId ThousandsSeparator $thousandsSep
    set style        "GermanStyle"
} elseif { [lsearch -nocase $argv "--English"] >= 0 } {
    $appId UseSystemSeparators false
    set floatSep     "." 
    set thousandsSep ","
    $appId DecimalSeparator   $floatSep
    $appId ThousandsSeparator $thousandsSep
    set style        "EnglishStyle"
} else {
    set floatSep     [Excel GetDecimalSeparator $appId]
    set thousandsSep [Excel GetThousandsSeparator $appId]
    set style        "NativeStyle"
}

puts "Using system separators  : [$appId UseSystemSeparators]"
puts "Using decimal   separator: [Excel GetDecimalSeparator $appId]"
puts "Using thousands separator: [Excel GetThousandsSeparator $appId]"

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-24_Format_$style"]
append xlsFile [Excel GetExtString $appId]
file delete -force $xlsFile

set worksheetId [Excel GetWorksheetIdByIndex $workbookId 1]
Excel SetWorksheetName $worksheetId $style

Excel SetHeaderRow $worksheetId [list "Value as text" "Number format" "Display"]

set row 2

set numberFormat [Excel GetNumberFormat $appId "0" "0000" $floatSep]
InsertValue $worksheetId $row "0.0"     $numberFormat  ; incr row
InsertValue $worksheetId $row "0.12345" $numberFormat  ; incr row
InsertValue $worksheetId $row "12345.6" $numberFormat  ; incr row

set numberFormat [Excel GetNumberFormat $appId "0" "0" $floatSep]
InsertValue $worksheetId $row "0.0"     $numberFormat  ; incr row
InsertValue $worksheetId $row "0.12345" $numberFormat  ; incr row
InsertValue $worksheetId $row "12345.6" $numberFormat  ; incr row

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
