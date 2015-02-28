# Test CawtExcel procedures for inserting images using cell background colors.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt
package require Tk

set testImg [file join [pwd] "testIn/wish.gif"]

# Open Excel, show the application window and create a workbook.
set appId [::Excel::Open true]
set workbookId [::Excel::AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-21_ImgCell"]
append xlsFile [::Excel::GetExtString $appId]
file delete -force $xlsFile

frame .f1
frame .f2
pack .f1 .f2 -side left

pack [label .f1.l] [label .f1.t] -side top
pack [label .f2.l] [label .f2.t] -side top


set phImg [image create photo -file $testImg]
set w [image width $phImg]
set h [image height $phImg]
set numPix [expr { $w * $h }]

.f1.l configure -image $phImg
.f1.t configure -text  "Original image (Size: $w x $h)"
update

puts "Add an image as cell background colors (at specific position) ..."
set worksheetId2 [::Excel::AddWorksheet $workbookId "ImageParam"]
::Excel::UseImgTransparency false

set t1 [clock clicks -milliseconds]
::Excel::ImgToWorksheet $phImg $worksheetId2 10 3  5 1
set t2 [clock clicks -milliseconds]
puts "[expr $t2 - $t1] ms to put $numPix pixels (ignoring transparency) into worksheet."

puts "Add an image as cell background colors (without screen update) ..."
set worksheetId3 [::Excel::AddWorksheet $workbookId "ImageNoUpdate"]
::Excel::UseImgTransparency false

::Excel::ScreenUpdate $appId false
set t1 [clock clicks -milliseconds]
::Excel::ImgToWorksheet $phImg $worksheetId3 10 3  5 1
set t2 [clock clicks -milliseconds]
puts "[expr $t2 - $t1] ms to put $numPix pixels (ignoring transparency) into worksheet."
::Excel::ScreenUpdate $appId true

puts "Add an image as cell background colors (Default params) ..."
set worksheetId4 [::Excel::AddWorksheet $workbookId "ImageDefault"]
::Excel::UseImgTransparency true

set t1 [clock clicks -milliseconds]
::Excel::ImgToWorksheet $phImg $worksheetId4
set t2 [clock clicks -milliseconds]
puts "[expr $t2 - $t1] ms to put $numPix pixels (using transparency) into worksheet."

puts "Paint a cross ..."
set rangeId [::Excel::SelectRangeByIndex $worksheetId4 [expr $h/2] 1 [expr $h/2 +1] $w]
::Excel::SetRangeFillColor $rangeId 255 0 255

set rangeId [::Excel::SelectRangeByIndex $worksheetId4 1 [expr $w/2] $h [expr $w/2 +1]]

::Excel::SetRangeFillColor $rangeId 255 0 255

puts "Export changed image into a Tk photo ..."
set t1 [clock clicks -milliseconds]
set phImgPainted [::Excel::WorksheetToImg $worksheetId4 1 1 $h $w]
set t2 [clock clicks -milliseconds]
puts "[expr $t2 - $t1] ms to get $numPix pixels from worksheet."

set wp [image width $phImgPainted]
set hp [image height $phImgPainted]

.f2.l configure -image $phImgPainted
.f2.t configure -text  "Painted image (Size: $wp x $hp)"
update

::Cawt::CheckNumber $w $wp "Width of images"
::Cawt::CheckNumber $h $hp "Height of images"

set imgFile [file join [pwd] "testOut" "Excel-21_ImgCell.gif"]
puts "Saving painted images: $imgFile"
$phImgPainted write $imgFile -format "GIF"

puts "Saving as Excel file: $xlsFile"
::Excel::SaveAs $workbookId $xlsFile "" false

::Cawt::PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
