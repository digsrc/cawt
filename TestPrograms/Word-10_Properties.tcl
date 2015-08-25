# Test CawtWord procedures related to property handling.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

# Open Word, show the application window and create a workbook.
set appId [Word OpenNew true]
set docId [Word AddDocument $appId]

# Delete Word file from previous test run.
file mkdir testOut
set docFile [file join [pwd] "testOut" "Word-10_Properties"]
append docFile [Word GetExtString $appId]
file delete -force $docFile

# Set some builtin and custom properties and check their values.
Cawt SetDocumentProperty $docId "Author"     "Paul Obermeier"
Cawt SetDocumentProperty $docId "Company"    "poSoft"
Cawt SetDocumentProperty $docId "Title"      $docFile
Cawt SetDocumentProperty $docId "CustomProp" "CustomValue"

Cawt CheckString "Paul Obermeier" [Cawt GetDocumentProperty $docId "Author"]     "Property Author"
Cawt CheckString "poSoft"         [Cawt GetDocumentProperty $docId "Company"]    "Property Company"
Cawt CheckString $docFile         [Cawt GetDocumentProperty $docId "Title"]      "Property Title"
Cawt CheckString "CustomValue"    [Cawt GetDocumentProperty $docId "CustomProp"] "Property CustomProp"


# Get all builtin and custom properties and insert them into the document.
Word AppendText $docId "Builtin Properties:"
set builtinProps [Cawt GetDocumentProperties $docId "Builtin"]
set builtinTable [Word AddTable [Word GetEndRange $docId] [expr [llength $builtinProps] +1] 2]
Word SetTableBorderLineStyle $builtinTable
Word SetHeaderRow $builtinTable [list "Name" "Value"]

set row 2
set col 1
foreach propertyName $builtinProps {
    Word SetCellValue $builtinTable $row [expr $col + 0] $propertyName
    Word SetCellValue $builtinTable $row [expr $col + 1] [Cawt GetDocumentProperty $docId $propertyName]
    incr row
}

Word AppendText $docId "Custom Properties:"
set customProps [Cawt GetDocumentProperties $docId "Custom"]
set customTable [Word AddTable [Word GetEndRange $docId] [expr [llength $customProps] +1] 2]
Word SetTableBorderLineStyle $customTable
Word SetHeaderRow $customTable [list "Name" "Value"]
set row 2
set col 1

foreach propertyName [Cawt GetDocumentProperties $docId "Custom"] {
    Word SetCellValue $customTable $row [expr $col + 0] $propertyName
    Word SetCellValue $customTable $row [expr $col + 1] [Cawt GetDocumentProperty $docId $propertyName]
    incr row
}

puts "Saving as Word file: $docFile"
Word SaveAs $docId $docFile

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Word Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
