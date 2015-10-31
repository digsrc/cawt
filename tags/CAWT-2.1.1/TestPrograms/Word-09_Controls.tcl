# Test CawtWord procedures for handling content controls.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}
package require cawt

# Open new Word instance and show the application window.
set appId [Word OpenNew true]

# Delete Word file from previous test run.
file mkdir testOut
set rootName [file join [pwd] "testOut" "Word-09_Controls"]
set wordFile [format "%s%s" $rootName [Word GetExtString $appId]]
file delete -force $wordFile

# Create a new document.
set docId [Word AddDocument $appId]

set tableId [Word AddTable [Word GetEndRange $docId] 10 2]
Word SetTableBorderLineStyle $tableId
Word SetHeaderRow $tableId [list "Content Control Type" "Example"]
set row 1

incr row
Word SetCellValue $tableId $row 1 "wdContentControlCheckBox"
set cellId [Word GetCellRange $tableId $row 2]
# Use in a catch statement, as content controls are available only in Word 2007 an up.
set catchVal [catch { Word AddContentControl $cellId wdContentControlCheckBox "CheckBox" } retVal]
if { $catchVal } {
    puts "Error: $retVal"
} else {
    incr row
    Word SetCellValue $tableId $row 1 "wdContentControlText"
    set cellId [Word GetCellRange $tableId $row 2]
    set controlId [Word AddContentControl $cellId wdContentControlText "Text"]
    Word SetContentControlText $controlId "What's your favorite language ?"

    incr row
    Word SetCellValue $tableId $row 1 "wdContentControlDropdownList"
    set cellId [Word GetCellRange $tableId $row 2]
    set controlId [Word AddContentControl $cellId wdContentControlDropdownList "Dropdown"]
    Word SetContentControlDropdown $controlId "Choose favorite language" [list Tcl Yes Python No]
}

# Save document as Word file.
puts "Saving as Word file: $wordFile"
Word SaveAs $docId $wordFile

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Word Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
