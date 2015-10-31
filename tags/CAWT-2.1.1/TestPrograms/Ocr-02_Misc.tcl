# Test miscellaneous CawtOcr procedures.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}
package require cawt

set inFile [file join [pwd] "testIn/ocr.bmp"]

set retVal [catch { package require Img } version]
if { $retVal == 0 } {
    pack [label .l]
    set phImg [image create photo -file $inFile]
    .l configure -image $phImg
    update
} else {
    puts "Warning: Img extension missing"
}

# Do the actual OCR scanning and store the scanned text in variable page.
set ocrId [Ocr Open]
Ocr OpenDocument $ocrId $inFile
puts "Number of images in input file: [Ocr GetNumImages $ocrId]"
set ocrLayout [Ocr Scan $ocrId 0]
set page [Ocr GetFullText $ocrLayout]

# Open new Word instance and create a new document.
set appWordId [Word OpenNew true]
set docId [Word AddDocument $appWordId]

# Open Excel, show the application window and create a workbook.
set appExcelId [Excel Open true]
set workbookId [Excel AddWorkbook $appExcelId]

# Delete test files from previous test run.
file mkdir testOut
set xlsFile [file join [pwd] "testOut" "Ocr-02_Misc"]
append xlsFile [Excel GetExtString $appExcelId]
file delete -force $xlsFile
set wordFile [file join [pwd] "testOut" "Ocr-02_Misc"]
append wordFile [Word GetExtString $appWordId]
file delete -force $wordFile

# Insert the recognized text into the Word document.
Word AppendText $docId $page
puts "Saving as Word file: $wordFile"
Word SaveAs $docId $wordFile

# Now get the recognition statistics from the OCR module and store
# it in an Excel worksheet.

# Select the first - already existing - worksheet,
# set its name and fill the header rows.
set worksheetId [Excel GetWorksheetIdByIndex $workbookId 1]
Excel SetWorksheetName $worksheetId "OcrMisc"
Excel SetHeaderRow $worksheetId \
    { "Text" "Id" "Line" "RegionId" "FontId" "Confidence" }

set numWords [Ocr GetNumWords $ocrLayout]
puts "Number of words recognized: $numWords"

set row 2
for { set i 0 } { $i < $numWords } { incr i } {
    set wordText  [Ocr GetWord $ocrLayout $i]
    set wordStats [Ocr GetWordStats $ocrLayout $i]
    Excel SetRowValues $worksheetId $row [list \
        $wordText \
        [dict get $wordStats Id] \
        [dict get $wordStats LineId] \
        [dict get $wordStats RegionId] \
        [dict get $wordStats FontId] \
        [dict get $wordStats Confidence] ]
    incr row
}
Excel SetColumnsWidth $worksheetId 1 6 0

Ocr Close $ocrId

puts "Saving as Excel file: $xlsFile"
Excel SaveAs $workbookId $xlsFile

if { [lindex $argv 0] eq "auto" } {
    Word Quit $appWordId
    Excel Quit $appExcelId
    Cawt Destroy
    exit 0
}
Cawt Destroy
