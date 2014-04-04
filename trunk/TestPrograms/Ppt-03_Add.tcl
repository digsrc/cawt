# Test CawtPpt procedures for adding and copying slides.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

set appId [::Ppt::Open]

# Delete PowerPoint files from previous test run.
file mkdir testOut
set pptFile1 [file join [pwd] "testOut" "Ppt-03_Add-1"]
append pptFile1 [::Ppt::GetExtString $appId]
file delete -force $pptFile1
set pptFile2 [file join [pwd] "testOut" "Ppt-03_Add-2"]
append pptFile2 [::Ppt::GetExtString $appId]
file delete -force $pptFile2

# Add 2 presentations to test the CloseAll method.
set presId1 [::Ppt::AddPres $appId]
puts "Active presentation: [::Ppt::GetActivePres $appId]"
set presId2 [::Ppt::AddPres $appId]
puts "Active presentation: [::Ppt::GetActivePres $appId]"

set imgName [file join [pwd] "testIn" "wish.gif"]

# Add a slide to each presentation and load the Wish image in different sizes.
set slideId1 [::Ppt::AddSlide $presId1]
set slideId2 [::Ppt::AddSlide $presId2]

set imgId1 [::Ppt::InsertImage $slideId1 $imgName \
           [::Cawt::CentiMetersToPoints 1] [::Cawt::CentiMetersToPoints 2] \
           [::Cawt::CentiMetersToPoints 6] [::Cawt::CentiMetersToPoints 6]]
set imgId2 [::Ppt::InsertImage $slideId2 $imgName \
           [::Cawt::CentiMetersToPoints 1] [::Cawt::CentiMetersToPoints 2]]

# Copy slide 1 to the end of the presentation.
set copiedSlide [::Ppt::CopySlide $presId1 1]

# Save both presentations.
puts "Saving as PowerPoint file: $pptFile1"
::Ppt::SaveAs $presId1 $pptFile1
puts "Saving as PowerPoint file: $pptFile2"
::Ppt::SaveAs $presId2 $pptFile2

# Close all open presentations.
::Ppt::CloseAll $appId

# Reopen presentation 1.
::Ppt::OpenPres $appId $pptFile1

if { [lindex $argv 0] eq "auto" } {
    ::Ppt::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
