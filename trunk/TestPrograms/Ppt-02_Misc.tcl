# Test miscellaneous CawtPpt procedures like adding slides, inserting images and saving slides
# as image files.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}
package require cawt

set appId [Ppt Open]
set presId [Ppt AddPres $appId]

# Delete PowerPoint file from previous test run.
file mkdir testOut
set pptFile [file join [pwd] "testOut" "Ppt-02_Misc"]
append pptFile [Ppt GetExtString $appId]
file delete -force $pptFile
set imgDir [file join [pwd] "testOut" "Ppt-02_Misc_Gif"]
file delete -force $imgDir

set imgName [file join [pwd] "testIn" "wish.gif"]

set slideId1 [Ppt AddSlide $presId]
set slideId2 [Ppt AddSlide $presId]
set slideId3 [Ppt AddSlide $presId]

set img1Id [Ppt InsertImage $slideId1 $imgName 1c 2c]
set img2Id [Ppt InsertImage $slideId2 $imgName 1c 2c 3c 3c]
set img3Id [Ppt InsertImage $slideId3 $imgName 1c 2c 6c 6c]

# Test switching the ViewType.
Ppt SetViewType $presId ppViewSlide
Cawt CheckNumber $Ppt::ppViewSlide [Ppt GetViewType $presId] "ViewType"

Ppt SetViewType $presId ppViewSlideSorter
Cawt CheckNumber $Ppt::ppViewSlideSorter [Ppt GetViewType $presId] "ViewType"

puts "Saving as PowerPoint file: $pptFile"
Ppt SaveAs $presId $pptFile

puts "Saving as GIF image files: $imgDir"
Ppt SaveAs $presId $imgDir ppSaveAsGIF

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Ppt Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
