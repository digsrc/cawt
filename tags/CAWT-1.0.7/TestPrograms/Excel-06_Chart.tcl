# Test CawtExcel procedures for creating charts and exporting charts as Tk photo images.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

# Name of Excel file being generated.
# Number of test rows and columns being generated.
# Note:
#   The actual number of rows will be numRows+1,
#   because the first row contains header information.
#   The actual number of columns will be numCols+1,
#   because the first column contains time information.
set xlsFile [file join [pwd] "testOut" "Excel-06_Chart"]
set numRows  10
set numCols   3

proc InsertTestData { worksheetId timeHeader timeList valsHeaderList valsList } {
    # Insert first row with header names.
    # First column contains the time header, next columns of row 1
    # contain the header of the data columns.

    ::Excel::SetCellValue $worksheetId 1 1 $timeHeader

    set col 2
    foreach head $valsHeaderList {
        ::Excel::SetCellValue $worksheetId 1 $col $head
        incr col
    }
    set lastDataCol [expr $col - 1]

    # Format the header lines.
    ::Excel::FormatHeaderRow $worksheetId 1  1 [expr $col-1]

    # Now insert the data row by row.
    set row 2
    foreach t $timeList vals $valsList {
        set col 1
        ::Excel::SetCellValue $worksheetId $row $col $t "real"
        incr col
        foreach val $vals {
            ::Excel::SetCellValue $worksheetId $row $col $val "real"
            incr col
        }
        incr row
    }
    set lastDataRow [expr $row - 1]

    incr row
    set r $row
    set funcList [list "MIN" "MAX" "AVERAGE" "STDEV"]
    foreach labelStr $funcList {
        ::Excel::SetCellValue $worksheetId $r 1 $labelStr
        set cellId [::Excel::GetCellIdByIndex $worksheetId $r 1]
        ::Excel::SetRangeFontBold $cellId true
        ::Cawt::Destroy $cellId
        incr r
    }

    for { set c 2 } { $c <= $lastDataCol } { incr c } {
        set dataRange [::Excel::GetCellRange 2 $c $lastDataRow $c]
        set r $row
        foreach func $funcList {
            set cellId [::Excel::SelectRangeByIndex $worksheetId $r $c $r $c]
            set formula [format "=%s(%s)" $func $dataRange]
            $cellId Formula $formula
            ::Cawt::Destroy $cellId
            incr r
        }
    }
}

# Test preparation: Generate lists with test data.

# The name of the header of column 1.
set timeHeader "Time"

# The names of the header of columns [2, numCols+1].
for { set c 1 } { $c <= $numCols } {incr c } {
    lappend valsHeaderList "Coord-$c"
}

# The values for the time column (1).
for { set r 0 } { $r < $numRows } {incr r } {
    lappend timeList [expr $r * 0.1]
}

# The values for the data columns [2, numCols+1].
set minVal  1.0E6
set maxVal -1.0E6
for { set r 1 } { $r <= $numRows } {incr r } {
    set colList [list]
    for { set c 1 } { $c <= $numCols } {incr c } {
        set newVal [expr {10 + $r +$c*0.6}]
        lappend colList $newVal
        if { $newVal > $maxVal } {
            set maxVal $newVal
        }
        if { $newVal < $minVal } {
            set minVal $newVal
        }
    }
    lappend valsList $colList
}

# Data for test 3 and 4: Create a RadarMark and ClusteredColumn chart.
expr srand (1)
for { set i 1 } { $i <= $numRows } { incr i } {
    lappend entList "Entity-$i"
    set shoots [expr int (20 * rand())]
    if { $shoots < 5 } {
        incr shoots 5
    }
    set hits [expr $shoots - int (10 * rand())]
    if { $hits < 0 } {
        set hits $shoots
    }
    if { $hits > $shoots } {
        set hits $shoots
    }
    lappend shootList $shoots
    lappend hitList $hits
}

# Test start: Open new Excel instance,
# show the application window and create a workbook.
set appId [::Excel::OpenNew]
set workbookId [::Excel::AddWorkbook $appId]

# Delete Excel file from previous test run.
append xlsFile [::Excel::GetExtString $appId]
file mkdir testOut
catch { file delete -force $xlsFile }

#
# Perform test 1:
# Interpret the data as flight paths and display each data column as a line.
# The time column is used for the X axis.

# Create a worksheet and set its name.
# We use the first already existing worksheet for our first test.
# Mainly because otherwise our charts are not placed on the intended
# worksheet, but on the first default one. (Bug in Excel 2000, in Excel 2003
# this works correctly.
# set worksheetId [::Excel::AddWorksheet $workbookId]
set worksheetId [::Excel::GetWorksheetIdByIndex $workbookId 1]
::Excel::SetWorksheetName $worksheetId "LineChart"

# Insert the list data into the Excel worksheet and automatically fit
# the column width.
InsertTestData $worksheetId $timeHeader $timeList $valsHeaderList $valsList

# Generate the line chart.
set lineChartId1 [::Excel::AddLineChartSimple $worksheetId \
                  $numRows $numCols "All flight paths"]
::Excel::SetChartMinScale $lineChartId1 "y" $minVal
::Excel::SetChartMaxScale $lineChartId1 "y" $maxVal
set lineChartObjId1 [::Excel::PlaceChart $lineChartId1 $worksheetId]
::Excel::SetChartObjPosition $lineChartObjId1 300 20

# AddLineChart worksheetId headerRow xaxisCol
#              startRow numRows startCol numCols
#              title yaxisName markerSize
set lineChartId2 [::Excel::AddLineChart $worksheetId \
                  1 1  3 4  3 2  "Some flight paths" "Coordinate"]
set lineChartObjId2 [::Excel::PlaceChart $lineChartId2 $worksheetId]
::Excel::SetChartObjPosition $lineChartObjId2 300 220

# Perform test 2:
# Interpret the data of columns 2 and 4 as a 2D location (lat, lon).
# and display the locations as a point chart.

# Create a worksheet and set its name.
set worksheetId [::Excel::AddWorksheet $workbookId "PointChart"]

# Insert the list data into the Excel worksheet.
InsertTestData $worksheetId $timeHeader $timeList $valsHeaderList $valsList

set pointChartId [::Excel::AddPointChartSimple $worksheetId \
                  $numRows 2 4 "MunitionDetonations"]
::Excel::SetChartScale $pointChartId $minVal $maxVal $minVal $maxVal
set pointChartObjId [::Excel::PlaceChart $pointChartId $worksheetId]
set chartRangeId [::Excel::SelectRangeByString $worksheetId "F2:K16"]
::Excel::ResizeChartObj $pointChartObjId $chartRangeId

# Perform test 3:
# Load data from entList, hitList and shootList into a worksheet.
# Display the data as a radar mark chart.

# Create a worksheet and set its name.
set worksheetId [::Excel::AddWorksheet $workbookId "RadarChart"]

# Insert the list data into the Excel worksheet.
::Excel::SetHeaderRow $worksheetId [list "Entity" "Shots" "Hits"]

::Excel::SetColumnValues $worksheetId 1 $entList   2
::Excel::SetColumnValues $worksheetId 2 $shootList 2
::Excel::SetColumnValues $worksheetId 3 $hitList   2

# Fit the column width automatically.
::Excel::SetColumnsWidth $worksheetId 1 3 0

set radarChartId [::Excel::AddRadarChartSimple $worksheetId $numRows 2]

# Place the radar chart as an object in the current worksheet.
set radarChartObjId [::Excel::PlaceChart $radarChartId $worksheetId]

# Set the size of the generated chart.
::Excel::SetChartObjSize $radarChartObjId 640 480

# Copy the radar chart to the Windows clipboard.
::Excel::ChartObjToClipboard $radarChartObjId

# Save the radar chart as a GIF file.
::Excel::SaveChartObjAsImage $radarChartObjId [file join [pwd] "testOut" "Excel-06_Chart.gif"]

# Perform test 4:
# Load data from entList, hitList and shootList into a worksheet.
# Display the data as a clustered column chart.

# Create a worksheet and set its name.
set worksheetId [::Excel::AddWorksheet $workbookId "ColumnChart"]

# Insert the list data into the Excel worksheet.
::Excel::SetHeaderRow $worksheetId [list "Entity" "Shots" "Hits"]

::Excel::SetColumnValues $worksheetId 1 $entList   2
::Excel::SetColumnValues $worksheetId 2 $shootList 2
::Excel::SetColumnValues $worksheetId 3 $hitList   2

# Fit the column width automatically.
::Excel::SetColumnsWidth $worksheetId 1 3 0

set columnChartId [::Excel::AddColumnChartSimple $worksheetId $numRows 2 "Clustered Column"]

# Place the column chart as an object in the current worksheet.
set columnChartObjId [::Excel::PlaceChart $columnChartId $worksheetId]

# Set the size of the placed chart object.
::Excel::SetChartObjSize $columnChartObjId 640 480

# Check number of rows in different range objects.
puts "Number of rows in worksheet   : [::Excel::GetNumRows    $worksheetId]"
puts "Number of columns in worksheet: [::Excel::GetNumColumns $worksheetId]"

set rangeId [::Excel::SelectRangeByIndex $worksheetId 2 1 \
                                         [expr $numRows+1] 3 true]
::Cawt::CheckNumber 10 [::Excel::GetNumRows    $rangeId] "Number of rows in range"
::Cawt::CheckNumber  3 [::Excel::GetNumColumns $rangeId] "Number of columns in range"

# Enable the auto filter menus.
::Excel::ToggleAutoFilter $rangeId

# If we have the Img and Twapi extension, get the chart as a photo image
# from the clipboard and create a Tk label to display it.
# Then resize the photo, copy the scaled image to the clipboard and paste it
# into a new worksheet into a specified cell.
set retVal [catch {::Cawt::ClipboardToImg} phImg]
if { $retVal == 0 } {
    label .l1
    label .l2
    pack .l1 .l2 -side left

    set width  [image width $phImg]
    set height [image height $phImg]
    set widthHalf  [expr {$width  / 2}]
    set heightHalf [expr {$height / 2}]

    set phImgHalf [image create photo -width $widthHalf -height $heightHalf]
    $phImgHalf copy $phImg -subsample 2 2

    wm title . "Extracted Excel chart as photo image (Size: $width x $height)"
    .l1 configure -image $phImg
    .l2 configure -image $phImgHalf

    set retVal [catch {::Cawt::ImgToClipboard $phImgHalf}]
    if { $retVal == 0 } {
        set pasteWorksheetId [::Excel::AddWorksheet $workbookId "ImagePaste"]
        set row 5
        set col 2
        set cellId [::Excel::SelectCellByIndex $pasteWorksheetId $row $col true]
        $pasteWorksheetId Paste
    } else {
        puts "Warning: Base64 extension missing"
    }
} else {
    puts "Warning: Img extension missing"
}

puts "Saving as Excel file: $xlsFile"
::Excel::SaveAs $workbookId $xlsFile

if { [lindex $argv 0] eq "auto" } {
    ::Excel::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
