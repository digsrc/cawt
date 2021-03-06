# Test CawtPpt procedures for adding and copying slides.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}
package require cawt

set appId [Ppt Open]

# Delete PowerPoint files from previous test run.
file mkdir testOut
set pptFile1 [file join [pwd] "testOut" "Ppt-03_Add-1"]
append pptFile1 [Ppt GetExtString $appId]
file delete -force $pptFile1
set pptFile2 [file join [pwd] "testOut" "Ppt-03_Add-2"]
append pptFile2 [Ppt GetExtString $appId]
file delete -force $pptFile2

# Add 2 presentations to test the CloseAll method.
set presId1 [Ppt AddPres $appId]
puts "Active presentation: [Ppt GetActivePres $appId]"
set presId2 [Ppt AddPres $appId]
puts "Active presentation: [Ppt GetActivePres $appId]"

set imgName [file join [pwd] "testIn" "wish.gif"]

# Add a slide to each presentation and load the Wish image in different sizes.
set slideId1 [Ppt AddSlide $presId1]
set slideId2 [Ppt AddSlide $presId2]

set imgId1 [Ppt InsertImage $slideId1 $imgName 1c 2c 6c 6c]
set imgId2 [Ppt InsertImage $slideId2 $imgName 1c 2c]

# Copy slide 1 of presId1 to the end of the presentation.
set copiedSlide1 [Ppt CopySlide $presId1 1]

# Copy slide 1 of presId1 to the end of presentation presId2.
set copiedSlide2 [Ppt CopySlide $presId1 1 end $presId2]

# Copy slide 1 of presId1 to the beginning of presentation presId2.
set copiedSlide2 [Ppt CopySlide $presId1 1 1 $presId2]

# Save both presentations.
puts "Saving as PowerPoint file: $pptFile1"
Ppt SaveAs $presId1 $pptFile1
puts "Saving as PowerPoint file: $pptFile2"
Ppt SaveAs $presId2 $pptFile2

# Close all open presentations.
Ppt CloseAll $appId

# Reopen presentation 2.
Ppt OpenPres $appId $pptFile2

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Ppt Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
