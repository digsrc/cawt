# Test CawtPpt procedures for using PowerPoint custom layouts.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

set appId [Ppt Open]

set tmplFile [file join [pwd] "testIn" "CustomLayout"]
append tmplFile [Ppt GetTemplateExtString $appId]

# Delete export files from previous test run.
file mkdir testOut
set pptFile [file join [pwd] "testOut" "Ppt-06_CustomLayout"]
append pptFile [Ppt GetExtString $appId]
file delete -force $pptFile

set presId [Ppt AddPres $appId $tmplFile]
puts "Add slides from template file $tmplFile"

set numLayouts [Ppt GetNumCustomLayouts $presId]
Cawt CheckNumber 4 $numLayouts "Number of layouts"

for { set layoutNum 1 } { $layoutNum <= $numLayouts } { incr layoutNum } {
    set customLayoutId [Ppt GetCustomLayoutId $presId $layoutNum]
    set layoutName [Ppt GetCustomLayoutName $customLayoutId]
    puts "  Adding $layoutName"
    lappend layoutNameList $layoutName
    Ppt AddSlide $presId $customLayoutId
}

set layoutName "Layout 3 - 2 widgets"
set customLayoutId [Ppt GetCustomLayoutId $presId $layoutName]
Ppt AddSlide $presId $customLayoutId
Cawt CheckString $layoutName [Ppt GetCustomLayoutName $customLayoutId] "Adding layout by name"

set customLayoutId [Ppt GetCustomLayoutId $presId end]
Ppt AddSlide $presId $customLayoutId
Cawt CheckString [lindex $layoutNameList end] [Ppt GetCustomLayoutName $customLayoutId] \
                    "Adding layout by special index end"

puts "Saving as PowerPoint file: $pptFile"
Ppt SaveAs $presId $pptFile

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Ppt Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
