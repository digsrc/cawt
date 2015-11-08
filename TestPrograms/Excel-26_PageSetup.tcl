# Test CawtExcel procedures related to page setup and printing.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}
package require cawt

# Number of test rows and columns being generated.
set numRows  10
set numCols   3

# Generate row list with test data
for { set i 1 } { $i <= $numCols } { incr i } {
    lappend rowList $i
}

# Open Excel, show the application window and create a workbook.
set appId [Excel Open true]
set workbookId [Excel AddWorkbook $appId]

# Delete Excel file from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Excel-26_PageSetup"]
append xlsFile [Excel GetExtString $appId]
file delete -force $xlsFile

set worksheetId(1) [Excel AddWorksheet $workbookId "PageSetup1"]
set worksheetId(2) [Excel AddWorksheet $workbookId "PageSetup2"]

for { set row 1 } { $row <= $numRows } { incr row } {
    Excel SetRowValues $worksheetId(1) $row $rowList
    Excel SetRowValues $worksheetId(2) $row $rowList
}

# Test the different page setup procedures.
proc PageSetup { worksheetId usePrinterComm } {
    set t1 [clock clicks -milliseconds]

    if { ! $usePrinterComm } {
        Cawt SetPrinterCommunication $worksheetId false
    }

    Excel SetWorksheetOrientation $worksheetId xlLandscape
    Excel SetWorksheetZoom        $worksheetId 50
    Excel SetWorksheetPaperSize   $worksheetId xlPaperA3
    Excel SetWorksheetFitToPages  $worksheetId 1 0

    Excel SetWorksheetPrintOptions $worksheetId \
          gridlines true \
          bw true \
          draft true \
          headings true \
          comments xlPrintSheetEnd \
          errors xlPrintErrorsNA

    # Set the header and footer texts.
    Excel SetWorksheetHeader $worksheetId \
          left   "LeftHeader"   \
          center "CenterHeader" \
          right  "RightHeader"
    Excel SetWorksheetFooter $worksheetId \
          left   "LeftFooter"   \
          center "CenterFooter" \
          right  "RightFooter"

    # Use a mixture of centimeters, inches and points for specifying the margin sizes.
    Excel SetWorksheetMargins $worksheetId \
          top 3c    bottom 2.5c \
          left 1i   right 2.1i  \
          footer 35 header 45p

    if { ! $usePrinterComm } {
        Cawt SetPrinterCommunication $worksheetId true
    }

    set t2 [clock clicks -milliseconds]
    if { $usePrinterComm } {
        puts "[expr $t2 - $t1] ms to set PageSetup properties using printer communication."
    } else {
        puts "[expr $t2 - $t1] ms to set PageSetup properties without printer communication."
    }
}

PageSetup $worksheetId(1) true
PageSetup $worksheetId(2) false

puts "Saving as Excel file: $xlsFile"
Excel SaveAs $workbookId $xlsFile "" false

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Excel Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
