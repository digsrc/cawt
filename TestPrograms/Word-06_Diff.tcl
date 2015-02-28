# Test CawtWord procedure for diff'ing Word files.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"

package require cawt

# Open Word, so we can get the extension string.
set appId [::Word::Open true]

set outPath [file join [pwd] "testOut"]
set wordBaseFile [file join $outPath Word-06_Diff-Base[::Word::GetExtString $appId]]
set wordNewFile  [file join $outPath Word-06_Diff-New[::Word::GetExtString $appId]]

# Create testOut directory, if it does not yet exist.
file mkdir testOut

# Delete Word output file from previous test run.
file delete -force $wordBaseFile
file delete -force $wordNewFile

for { set i 0 } { $i < 10 } { incr i } {
    append msg1 "This is line $i.\n"
}
for { set i 0 } { $i < 20 } { incr i } {
    append msg2 "This is line $i.\n"
}

# Create 2 Word files with some test data.
puts "Generating base file $wordBaseFile ..."
set docId [::Word::AddDocument $appId]
::Word::AppendText $docId "Base File" true
::Word::AppendText $docId $msg1 true
::Word::SaveAs $docId $wordBaseFile

puts "Generating new file $wordNewFile ..."
set docId [::Word::AddDocument $appId]
::Word::AppendText $docId "New File" true
::Word::AppendText $docId $msg2 true
::Word::SaveAs $docId $wordNewFile

::Word::Close $docId
::Word::Quit $appId

puts "Comparing base and new file ..."
set diffAppId [::Word::DiffWordFiles $wordBaseFile $wordNewFile]

::Cawt::PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    ::Word::Quit $diffAppId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
