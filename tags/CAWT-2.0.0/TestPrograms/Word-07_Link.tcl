# Test CawtWord procedures for handling links and inserting files.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

# Open new Word instance and show the application window.
set appId [Word OpenNew true]

# Delete Word file from previous test run.
file mkdir testOut
set wordFile [file join [pwd] "testOut" "Word-07_Link"]
append wordFile [Word GetExtString $appId]
file delete -force $wordFile

set inFile [file join [pwd] "testIn" "InsertMe.html"]
# set inFile [file join [pwd] ".." "Documentation" "UserManual" "CawtManualTemplate.doc"]

# Create a new document.
set docId [Word AddDocument $appId]

# Generate a text file for testing the external hyperlink capabilities.
set fileName [file join [pwd] "testOut" "Word-07_Link.txt"]
set fp [open $fileName "w"]
puts $fp "This is the text file linked from Word."
close $fp

set rangeLink [Word AppendText $docId "Dummy"]
Word SetHyperlink $rangeLink [format "file://%s" $fileName] "File Link"
Word AppendParagraph $docId

# Add internal links for the external file inserted later.
for { set i 1 } { $i <= 3 } { incr i } {
    Word AppendParagraph $docId
    set rangeLink [Word AppendText $docId "Dummy"]
    Word SetInternalHyperlink $rangeLink "proc$i" "Link to procedure $i"
}

Word AddPageBreak [Word GetEndRange $docId]

# Insert external file with different options.

puts "Insert external file via Word InsertFile method ..."
Word AppendText $docId "Inserted external file via Word InsertFile method" true
set endRange [Word GetEndRange $docId]
Word InsertFile $endRange $inFile

puts "Insert external file via PasteAndFormat wdPasteDefault ..."
Word AppendText $docId "Inserted external file via PasteAndFormat wdPasteDefault" true
set endRange [Word GetEndRange $docId]
Word InsertFile $endRange $inFile wdPasteDefault

puts "Insert external file via PasteAndFormat wdFormatOriginalFormatting ..."
Word AppendText $docId "Inserted external file via PasteAndFormat wdFormatOriginalFormatting" true
set endRange [Word GetEndRange $docId]
Word InsertFile $endRange $inFile wdFormatOriginalFormatting

# Save document as Word file.
puts "Saving as Word file: $wordFile"
Word SaveAs $docId $wordFile

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Word Quit $appId false
    Cawt Destroy
    exit 0
}
Cawt Destroy
