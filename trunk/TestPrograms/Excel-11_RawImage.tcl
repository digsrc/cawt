# Test CawtExcel procedures to exchange data between Excel and RAW photo images.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

set rawFile "testIn/gradient.raw"

set retVal [catch {package require img::raw} version]
if { $retVal == 0 } {
    set phImg [image create photo -file $rawFile \
               -format "RAW -useheader 1 -verbose 0 -nomap 0 -gamma 1"]
    pack [label .l] -side left
    .l configure -image $phImg
    wm title . "Original RAW image vs. generated RAW images \
               (Size: [image width $phImg] x [image height $phImg])"
    update
} else {
    puts "Package img::raw is not available."
    puts "Raw images will be read and written, but not displayed in a Tk window."
}

# Open new instance of Excel and add a workbook.
set excelAppId1 [::Excel::OpenNew]
set workbookId  [::Excel::AddWorkbook $excelAppId1]

# Delete Excel and RAW image files from previous test run.
file mkdir testOut
set xlsOutFile1 [file join [pwd] "testOut" "Excel-11_RawImage1"]
append xlsOutFile1 [::Excel::GetExtString $excelAppId1]
file delete -force $xlsOutFile1
set xlsOutFile2 [file join [pwd] "testOut" "Excel-11_RawImage2"]
append xlsOutFile2 [::Excel::GetExtString $excelAppId1]
file delete -force $xlsOutFile2
set rawOutFile1 [file join [pwd] "testOut" "Excel-11_RawImage1.raw"]
file delete -force $rawOutFile1
set rawOutFile2 [file join [pwd] "testOut" "Excel-11_RawImage2.raw"]
file delete -force $rawOutFile2
set rawOutFile3 [file join [pwd] "testOut" "Excel-11_RawImage3.raw"]
file delete -force $rawOutFile3

# Transfer image data with header information into Excel and vice versa.
set useHeader true

set worksheetId [::Excel::AddWorksheet $workbookId "WithHeader"]

set t1 [clock clicks -milliseconds]
::Excel::RawImageFileToWorksheet $rawFile $worksheetId $useHeader
set t2 [clock clicks -milliseconds]
puts "RawImageFileToWorksheet: [expr $t2 - $t1] ms (using header: $useHeader)."

set t1 [clock clicks -milliseconds]
::Excel::WorksheetToRawImageFile $worksheetId $rawOutFile1 $useHeader
set t2 [clock clicks -milliseconds]
puts "WorksheetToRawImageFile: [expr $t2 - $t1] ms (using header: $useHeader)."

# Transfer image data without header information into Excel and vice versa.
set useHeader false

set worksheetId [::Excel::AddWorksheet $workbookId "NoHeader"]

set t1 [clock clicks -milliseconds]
::Excel::RawImageFileToWorksheet $rawFile $worksheetId $useHeader
set t2 [clock clicks -milliseconds]
puts "RawImageFileToWorksheet: [expr $t2 - $t1] ms (using header: $useHeader)."

set t1 [clock clicks -milliseconds]
::Excel::WorksheetToRawImageFile $worksheetId $rawOutFile2 $useHeader
set t2 [clock clicks -milliseconds]
puts "WorksheetToRawImageFile: [expr $t2 - $t1] ms (using header: $useHeader)."

if { $retVal == 0 } {
    set phImg1 [image create photo -file $rawOutFile1 \
                -format "RAW -useheader 1 -verbose 0 -nomap 0 -gamma 1"]
    pack [label .l1] -side left
    .l1 configure -image $phImg1
    set phImg2 [image create photo -file $rawOutFile2 \
                -format "RAW -useheader 1 -verbose 0 -nomap 0 -gamma 1"]
    pack [label .l2] -side left
    .l2 configure -image $phImg2
}

puts "Saving as Excel file: $xlsOutFile1"
::Excel::SaveAs $workbookId $xlsOutFile1

puts "Convert raw image file $rawFile to Excel file."
set excelAppId2 [::Excel::RawImageFileToExcelFile $rawFile $xlsOutFile2 true false]

puts "Convert Excel file $xlsOutFile2 to raw image file."
::Excel::ExcelFileToRawImageFile $xlsOutFile2 $rawOutFile3 1 true true

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $excelAppId1
    ::Excel::Quit $excelAppId2
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
