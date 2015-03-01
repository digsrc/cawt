# Load position information into an Excel sheet, read back that information and create
# a Tk GUI with buttons corresponding to these positions.
# Clicking onto one of these buttons triggers Google Earth to fly to that position.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

file mkdir testOut

# Name of Excel file being generated.
set xlsFile [file join [pwd] "testOut" "Earth-02_MunichTour"]

# Name template of generated images.
set imgTemplate [file join [pwd] "testOut" "Earth-02_MunichTour-%d.jpg"]

# A (Google Earth) position is defined by the following 5 values.
set headerList { "Latitude" "Longitude" "Altitude" "Elevation" "Azimuth" }
set numCols [llength $headerList]

# Start (Munich center) and end (where I live) location.
set loc(0,lat) 48.1372
set loc(0,lon) 11.5753
set loc(0,alt) 700.0
set loc(0,ele) 55.0
set loc(0,azi) 15.0

set loc(1,lat) 48.1172
set loc(1,lon) 11.4828
set loc(1,alt) 500.0
set loc(1,ele) 70.0
set loc(1,azi) -90.0

set numRows 10

# Test start: Open new Excel instance,
# show the application window and create a workbook.
set appId [Excel OpenNew]
set workbookId [Excel AddWorkbook $appId]

# Delete Excel file from previous test run.
append xlsFile [Excel GetExtString $appId]
catch { file delete -force $xlsFile }

# Create a worksheet and set its name.
set worksheetId [Excel GetWorksheetIdByIndex $workbookId 1]
Excel SetWorksheetName $worksheetId "Munich Locations"

# Insert the header line.
Excel SetHeaderRow $worksheetId $headerList

set ind 0
set delta [expr $numRows - 1]
set dLat  [expr ($loc(1,lat) - $loc(0,lat)) / $delta]
set dLon  [expr ($loc(1,lon) - $loc(0,lon)) / $delta]
set dAlt  [expr ($loc(1,alt) - $loc(0,alt)) / $delta]
set dEle  [expr ($loc(1,ele) - $loc(0,ele)) / $delta]
set dAzi  [expr ($loc(1,azi) - $loc(0,azi)) / $delta]

set fmt1 [Excel GetLangNumberFormat "0" "0"]
set fmt2 [Excel GetLangNumberFormat "0" "00"]
set fmt4 [Excel GetLangNumberFormat "0" "0000"]

for { set row 2 } { $row <= [expr $numRows +1] } { incr row } {
    Excel SetCellValue $worksheetId $row 1 [expr $ind * $dLat + $loc(0,lat)] "real" $fmt4
    Excel SetCellValue $worksheetId $row 2 [expr $ind * $dLon + $loc(0,lon)] "real" $fmt4
    Excel SetCellValue $worksheetId $row 3 [expr $ind * $dAlt + $loc(0,alt)] "real" $fmt1
    Excel SetCellValue $worksheetId $row 4 [expr $ind * $dEle + $loc(0,ele)] "real" $fmt2
    Excel SetCellValue $worksheetId $row 5 [expr $ind * $dAzi + $loc(0,azi)] "int"
    incr ind
}

for { set row 2 } { $row <= [expr $numRows +1] } { incr row } {
    set lat [Excel GetCellValue $worksheetId $row 1 "real"]
    set lon [Excel GetCellValue $worksheetId $row 2 "real"]
    set alt [Excel GetCellValue $worksheetId $row 3 "real"]
    set ele [Excel GetCellValue $worksheetId $row 4 "real"]
    set azi [Excel GetCellValue $worksheetId $row 5 "int"]
    lappend posList [list $lat $lon $alt $ele $azi]
}

# Generate a point chart showing the different locations.
set chartId [Excel AddPointChartSimple $worksheetId $numRows 1 2 "Munich Tour"]
Excel PlaceChart $chartId $worksheetId

puts "Saving as Excel file: $xlsFile"
Excel SaveAs $workbookId $xlsFile

package require Tk

proc ShowLocation { geApp ind row } {
    global posList

    if { [Earth IsInitialized $geApp] } {
        set loc [lindex $posList $ind]
        set lat [lindex $loc 0]
        set lon [lindex $loc 1]
        set alt [lindex $loc 2]
        set ele [lindex $loc 3]
        set azi [lindex $loc 4]

        Earth SetCamera $geApp $lat $lon $alt $ele $azi
        set fileName [format $::imgTemplate $row]
        Earth SaveImage $geApp $fileName
    }
}

proc Quit { appId geApp } {
    Excel Quit $appId
    if { $geApp ne "" } {
        Earth Quit $geApp
    }
    ::Cawt::Destroy
    exit 0
}

puts "Starting GoogleEarth. This may take some time ..."
set geApp [Earth Open]
puts "Starting GUI ..."

if { $geApp ne "" } {
    pack [label .l -text "Munich tour with Google Earth controlled by Cawt"]
    set ind 0
    for { set row 2 } { $row <= [expr $numRows +1] } { incr row } {
        pack [button .b$row -text "Location (Row $row)" \
                            -command "ShowLocation $geApp $ind $row"] \
                            -expand true -fill x
        incr ind
    }
} else {
    pack [label .l -text "Google Earth not available"]
}
pack [button .b -text "Quit" -command "Quit $appId $geApp"] -expand true -fill x
wm title . "Munich Tour"

if { [lindex $argv 0] eq "auto" } {
    Quit $appId $geApp
    exit 0
}

