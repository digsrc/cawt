# Test CawtExcel procedures for dealing with images.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt
package require Tk

# Open Excel, show the application window and create a workbook.
set appId [::Excel::Open true]
set workbookId [::Excel::AddWorkbook $appId]

set worksheetId1 [::Excel::AddWorksheet $workbookId "Image"]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-20_ImgUtil"]
append xlsFile [::Excel::GetExtString $appId]
file delete -force $xlsFile

# Test inserting and scaling an image into a worksheet.
puts "Insert an image with Excel ..."
set picId [::Excel::InsertImage $worksheetId1 [file join [pwd] "testIn/wish.gif"] 1 1]
::Excel::ScaleImage $picId 2 2.5

set phImg [image create photo -file [file join [pwd] "testIn/wish.gif"]]
set w [image width $phImg]
set h [image height $phImg]

puts "Add an image as cell background colors (at specific position) ..."
set worksheetId2 [::Excel::AddWorksheet $workbookId "ImageParam"]
::Excel::ImgToWorksheet $phImg $worksheetId2 10 3  5 1

puts "Add an image as cell background colors (Default params) ..."
set worksheetId3 [::Excel::AddWorksheet $workbookId "ImageDef"]
::Excel::ImgToWorksheet $phImg $worksheetId3

set rangeId [::Excel::SelectRangeByIndex $worksheetId3 [expr $h/2] 1 [expr $h/2 +1] $w]
::Excel::SetRangeFillColor $rangeId 255 0 255
set rangeId [::Excel::SelectRangeByIndex $worksheetId3 1 [expr $w/2] $h [expr $w/2 +1]]
::Excel::SetRangeFillColor $rangeId 255 0 255

puts "Saving as Excel file: $xlsFile"
::Excel::SaveAs $workbookId $xlsFile "" false

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
