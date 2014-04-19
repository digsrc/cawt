# Test CawtExcel procedures to exchange data between Excel and Matlab files.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

set matFile [file join "testIn" "gradient.mat"]

# Open new instance of Excel and add a workbook.
set excelAppId1 [::Excel::OpenNew]
set workbookId  [::Excel::AddWorkbook $excelAppId1]

# Delete Excel and Matlab files from previous test run.
file mkdir testOut
set xlsOutFile1 [file join [pwd] "testOut" "Excel-12_MatlabFile1"]
append xlsOutFile1 [::Excel::GetExtString $excelAppId1]
file delete -force $xlsOutFile1
set xlsOutFile2 [file join [pwd] "testOut" "Excel-12_MatlabFile2"]
append xlsOutFile2 [::Excel::GetExtString $excelAppId1]
file delete -force $xlsOutFile2
set matOutFile1 [file join [pwd] "testOut" "Excel-12_MatlabFile1.mat"]
file delete -force $matOutFile1
set matOutFile2 [file join [pwd] "testOut" "Excel-12_MatlabFile2.mat"]
file delete -force $matOutFile2
set matOutFile3 [file join [pwd] "testOut" "Excel-12_MatlabFile3.mat"]
file delete -force $matOutFile3

# Transfer Matlab data with header information into Excel and vice versa.
set useHeader true

set worksheetId [::Excel::AddWorksheet $workbookId "WithHeader"]

set t1 [clock clicks -milliseconds]
::Excel::MatlabFileToWorksheet $matFile $worksheetId $useHeader
set t2 [clock clicks -milliseconds]
puts "MatlabFileToWorksheet: [expr $t2 - $t1] ms (using header: $useHeader)."

set t1 [clock clicks -milliseconds]
::Excel::WorksheetToMatlabFile $worksheetId $matOutFile1 $useHeader
set t2 [clock clicks -milliseconds]
puts "WorksheetToMatlabFile: [expr $t2 - $t1] ms (using header: $useHeader)."

# Transfer Matlab data without header information into Excel and vice versa.
set useHeader false

set worksheetId [::Excel::AddWorksheet $workbookId "NoHeader"]

set t1 [clock clicks -milliseconds]
::Excel::MatlabFileToWorksheet $matFile $worksheetId $useHeader
set t2 [clock clicks -milliseconds]
puts "MatlabFileToWorksheet: [expr $t2 - $t1] ms (using header: $useHeader)."

set t1 [clock clicks -milliseconds]
::Excel::WorksheetToMatlabFile $worksheetId $matOutFile2 $useHeader
set t2 [clock clicks -milliseconds]
puts "WorksheetToMatlabFile: [expr $t2 - $t1] ms (using header: $useHeader)."

puts "Saving as Excel file: $xlsOutFile1"
::Excel::SaveAs $workbookId $xlsOutFile1

puts "Convert Matlab file $matFile to Excel file."
set excelAppId2 [::Excel::MatlabFileToExcelFile $matFile $xlsOutFile2 true false]

puts "Convert Excel file $xlsOutFile2 to Matlab file."
::Excel::ExcelFileToMatlabFile $xlsOutFile2 $matOutFile3 1 true true

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $excelAppId1
    ::Excel::Quit $excelAppId2
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
