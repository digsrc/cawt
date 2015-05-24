# Test CawtWord procedures for handling content controls.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
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

set range [Word AppendText $docId "This is a check box (wdContentControlCheckBox): "]
Word AddContentControl [Word GetEndRange $docId] wdContentControlCheckBox
Word AppendParagraph $docId

set range [Word AppendText $docId "This is a text field (wdContentControlText): "]
Word AddContentControl [Word GetEndRange $docId] wdContentControlText
Word AppendParagraph $docId

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
