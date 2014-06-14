# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

set docDir   [pwd]
set cawtDir  [file join $docDir ".."]
set finalDir [file join $docDir "Final"]

set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

# Possible targets: all, ref, user, clean, distclean
set option "all"
if { $argc > 0 } {
    set option [lindex $argv 0]
}

if { $option eq "clean" || $option eq "distclean" } {
    file delete -force $finalDir
    exit 0
}

if { ! [file isdirectory $finalDir] } {
    file mkdir $finalDir
}

set retVal [catch {package require cawt} pkgVersion]

puts [format "%-25s: %s" "Tcl version" [info patchlevel]]
puts [format "%-25s: %s" "Cawt version" $pkgVersion]
puts [format "%-25s: %s" "Twapi version" [::Cawt::GetPkgVersion "twapi"]]

if { $option eq "ref" || $option eq "all" } {
    cd $cawtDir

    package require ruff

    puts "Generating reference documentation from source code ..."
    ::ruff::document_namespaces html \
        [list ::Cawt ::Word ::Excel ::Explorer ::Ppt ::Ocr ::Matlab ::Earth] \
        -title "CAWT Reference" \
        -output [file join $finalDir "CawtReference-$pkgVersion.html"]
    cd $docDir
}

if { $option eq "user" || $option eq "all" } {
    cd $docDir

    puts "Generating user manual from Word and PowerPoint files ..."
    set pptInFile  [file join [pwd] "UserManual" "CawtFigures.ppt"]
    set wordInFile [file join [pwd] "UserManual" "CawtManualTemplate.doc"]

    set userManFile [file join $finalDir "CawtManual-$pkgVersion.doc"]
    set pdfManFile  [file join $finalDir "CawtManual-$pkgVersion.pdf"]

    set outFigureDir [file join $finalDir "CawtFigures"]
    set testDir [file join $cawtDir "TestPrograms"]

    # Generate the figures for the user manual from the PowerPoint file.
    ::Ppt::ExportPptFile $pptInFile $outFigureDir "Figure-%02d.png" 1 end "PNG" -1 -1 false false

    # Copy the user manual template to new location and name.
    # Open the user manual template and perform the following actions:
    #   Insert a table with the test programs available (from folder TestPrograms).
    #   Insert the generated figures replacing the placeholder text.
    # Then save the finished user manual in the Final folder.
    file copy -force $wordInFile $userManFile
    set wordId [::Word::OpenNew]
    set docId [::Word::OpenDocument $wordId $userManFile false]
    ::Word::SetCompatibilityMode $docId $::Word::wdWord2003

    set placeHolder "%TABLE TestPrograms%"
    set numTables [::Word::GetNumTables $docId]
    set foundPlaceHolder false
    for { set n 1 } { $n <= $numTables } {incr n } {
        set tableId [::Word::GetTableIdByIndex $docId $n]
        # Placeholder must be listed in row 2, column 1.
        set cellCont [::Word::GetCellValue $tableId 2 1]
        if { $cellCont eq $placeHolder } {
            puts "    Replacing placeholder \"$placeHolder\" with list of test programs."
            set testFileList [lsort [glob -directory $testDir Earth* Excel* Explorer* Matlab* Ocr* Ppt* Word*]]
            set numRows [::Word::GetNumRows $tableId]
            set missingRows [expr [llength $testFileList] - $numRows +1]
            for { set r 1 } { $r <= $missingRows } {incr r } {
                ::Word::AddRow $tableId
            }

            set row 2
            foreach testFile $testFileList {
                set fp [open $testFile "r"]
                set description ""
                while { [gets $fp line] >= 0 } {
                    if { $line eq "#" } break
                    append description [string trim [string trim $line "#"]]
                    append description " "
                }
                close $fp
                set f [file tail $testFile]
                ::Word::SetCellValue $tableId $row 1 $f
                ::Word::SetCellValue $tableId $row 2 $description
                incr row
            }
            set foundPlaceHolder true
            break
        }
    }
    if { ! $foundPlaceHolder } {
        puts "    Warning: Placeholder \"$placeHolder\" not referenced in Word-Document"
    }

    set cropValues(Figure-01)  5.0    ; # Overview CAWT
    set cropValues(Figure-02)  9.0    ; # Module excelCsv
    set cropValues(Figure-03) 11.0    ; # Module excelTablelist
    set cropValues(Figure-04)  9.0    ; # Module excelMatlabFile
    set cropValues(Figure-05) 11.0    ; # Module excelWord
    set cropValues(Figure-06)  9.0    ; # Module excelImgRaw
    set cropValues(Figure-07)  9.0    ; # Module excelMediaWiki
    set cropValues(Figure-08)  9.0    ; # Module excelWikit

    foreach fig [glob -directory $outFigureDir *] {
        set figImg  [file tail $fig]
        set figName [file rootname $figImg]
        set placeHolder "%FIGURE $figName%"
        set myRange [::Word::GetStartRange $docId]
        if { ! [::Word::FindString $myRange $placeHolder] } {
            puts "    Warning: Figure $figName not referenced in Word-Document"
            continue
        } else {
            puts "    Replacing keyword $placeHolder with figure $figImg"
        }
        set rangeId [::Word::GetSelectionRange $docId]
        set imgId [::Word::InsertImage $rangeId $fig]
        if { [info exists cropValues($figName)] } {
            set crop [::Cawt::CentiMetersToPoints $cropValues($figName)]
        } else {
            puts "    Warning: No crop value specified for figure $figName"
            set crop 0.0
        }
        ::Word::CropImage $imgId $crop
    }
    ::Word::UpdateFields $docId
    
    ::Word::SaveAs $docId $userManFile
    set retVal [catch {::Word::SaveAsPdf $docId $pdfManFile} errMsg]
    if { $retVal } {
        puts "Warning: $errMsg"
    }
    ::Word::Close $docId
    ::Word::Quit $wordId
    ::Cawt::Destroy
}

exit 0
