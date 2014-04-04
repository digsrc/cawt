# Test CawtExcel procedures related to CSV files.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

set inPath  [file join [pwd] "testIn"]
set outPath [file join [pwd] "testOut"]

# Names of CSV files being generated.
set outFileExcel [file join $outPath Excel-07_Csv_Excel.csv]
set outFileCsv   [file join $outPath Excel-07_Csv_Csv.csv]

set inFileMediaWiki   [file join $inPath  MediaWikiTable.txt]
set outFileMediaWiki1 [file join $outPath Excel-07_Csv_MediaWiki1.txt]
set outFileMediaWiki2 [file join $outPath Excel-07_Csv_MediaWiki2.txt]

set inFileWikit   [file join $inPath  WikitTable.txt]
set outFileWikit1 [file join $outPath Excel-07_Csv_Wikit1.txt]
set outFileWikit2 [file join $outPath Excel-07_Csv_Wikit2.txt]

file mkdir testOut

# Add a workbook, add a worksheet and save it in CSV format.
set appId [::Excel::Open]
set workbookId  [::Excel::AddWorkbook $appId]
set worksheetId [::Excel::AddWorksheet $workbookId "CsvSep"]

# Insert some matrix data.
set testList {
    { 1 2 3 }
    { 4.1 5.2 6.2 }
    { 7,1 8,2 9,3 }
    { 3|1 4|2 5|3 }
    { "Hello; world" "What's" "next" }
}
::Excel::SetMatrixValues $worksheetId $testList
::Excel::SetMatrixValues $worksheetId $testList [expr [llength $testList] + 2]

puts "Saving CSV file $outFileExcel with Excel"
::Excel::SaveAsCsv $workbookId $worksheetId $outFileExcel
::Excel::Close $workbookId
::Excel::Quit $appId false

# Read the generated CSV file with the Cawt procedures and write it to a new CSV file.
::Excel::SetCsvSeparatorChar ","
::Cawt::CheckString "," [::Excel::GetCsvSeparatorChar] "::Excel::GetCsvSeparatorChar"
puts "Reading CSV file $outFileExcel"
set csvList [::Excel::ReadCsvFile $outFileExcel]
puts "Writing CSV file $outFileCsv"
::Excel::WriteCsvFile $csvList $outFileCsv

# Use the matrix generated above and write it to a new MediaWiki file.
puts "Writing MediaWiki file $outFileMediaWiki1"
::Excel::WriteMediaWikiFile $csvList $outFileMediaWiki1

# Read the MediaWiki test file (including potential column headers)
# and write it out again.
puts "Reading MediaWiki file $inFileMediaWiki"
set mediaWikiList [::Excel::ReadMediaWikiFile $inFileMediaWiki]
puts "Writing MediaWiki file $outFileMediaWiki2"
::Excel::WriteMediaWikiFile $mediaWikiList $outFileMediaWiki2

# Use the matrix generated above and write it to a new Wikit file.
puts "Writing Wikit file $outFileWikit1"
::Excel::WriteWikitFile $csvList $outFileWikit1

# Read the Wikit test file (including potential column headers)
# and write it out again.
puts "Reading Wikit file $inFileWikit"
set wikitList [::Excel::ReadWikitFile $inFileWikit]
puts "Writing Wikit file $outFileWikit2"
::Excel::WriteWikitFile $wikitList $outFileWikit2

::Cawt::Destroy
if { [lindex $argv 0] eq "auto" } {
    exit 0
}
