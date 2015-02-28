# Test CawtExcel procedures for dealing with images.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

set squareImg    [file join [pwd] "testIn/Square.gif"]
set landscapeImg [file join [pwd] "testIn/Landscape.gif"]
set portraitImg  [file join [pwd] "testIn/Portrait.gif"]

# Open Excel, show the application window and create a workbook.
set appId [::Excel::Open true]
set workbookId [::Excel::AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-20_ImgUtil"]
append xlsFile [::Excel::GetExtString $appId]
file delete -force $xlsFile

puts "Inserting images of different sizes ..."
set worksheetId1 [::Excel::AddWorksheet $workbookId "Sizes"]
::Excel::SetCellValue $worksheetId1 1 1 "Square"
::Excel::SetCellValue $worksheetId1 1 5 "Landscape"
::Excel::SetCellValue $worksheetId1 1 9 "Portrait"
::Excel::InsertImage $worksheetId1 $squareImg    2 1
::Excel::InsertImage $worksheetId1 $landscapeImg 2 5
::Excel::InsertImage $worksheetId1 $portraitImg  2 9

puts "Inserting images with different modes ..."
set worksheetId2 [::Excel::AddWorksheet $workbookId "Modes"]
::Excel::SetCellValue $worksheetId2 1 1 "Linked"
::Excel::SetCellValue $worksheetId2 1 5 "Embedded"
::Excel::SetCellValue $worksheetId2 1 9 "Linked and Embedded"
::Excel::InsertImage $worksheetId2 $squareImg  2 1  true  false
::Excel::InsertImage $worksheetId2 $squareImg  2 5  false true
::Excel::InsertImage $worksheetId2 $squareImg  2 9  true  true

set catchVal [ catch { ::Excel::InsertImage $worksheetId2 $squareImg 2 13 false false } retVal]
if { $catchVal } {
    puts "Error successfully caught: $retVal"
}

puts "Inserting and scaling images ..."
set worksheetId3 [::Excel::AddWorksheet $workbookId "Scaling"]
::Excel::SetCellValue $worksheetId3 1 1 "Landscape scaled to Square"
::Excel::SetCellValue $worksheetId3 1 5 "Portrait scaled to Square"
set shapeId1 [::Excel::InsertImage $worksheetId3 $landscapeImg 2 1]
::Excel::ScaleImage $shapeId1 1 2
set shapeId2 [::Excel::InsertImage $worksheetId3 $portraitImg 2 5]
::Excel::ScaleImage $shapeId2 2 1

puts "Saving as Excel file: $xlsFile"
::Excel::SaveAs $workbookId $xlsFile "" false

::Cawt::PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
