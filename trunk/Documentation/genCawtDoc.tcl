# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
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
        [list ::Cawt ::Word ::Excel ::Explorer ::Outlook ::Ppt ::Ocr ::Matlab ::Earth] \
        -title "CAWT Reference" \
        -output [file join $finalDir "CawtReference-$pkgVersion.html"]
    cd $docDir
}

if { $option eq "user" || $option eq "all" } {
    cd $docDir

    # For production these options must all be set to true.
    # Edit for testing purposes of a specific section.
    set optExportPpt           true
    set optReplaceRefTables    true
    set optReplaceTestPrograms true
    set optReplaceFigures      true
    set optReplaceKeywords     true
    set optHideChecks          true

    set printChecks [expr ! $optHideChecks]

    set refUrl "http://www.posoft.de/download/extensions/Cawt/CawtReference.html"

    puts "Generating user manual from Word and PowerPoint templates ..."
    set pptInFile  [file join [pwd] "UserManual" "CawtFigures.ppt"]
    set wordInFile [file join [pwd] "UserManual" "CawtManualTemplate.doc"]

    set userManFile [file join $finalDir "CawtManual-$pkgVersion.doc"]
    set pdfManFile  [file join $finalDir "CawtManual-$pkgVersion.pdf"]

    set outFigureDir [file join $finalDir "CawtFigures"]
    set testDir [file join $cawtDir "TestPrograms"]

    Cawt::CheckComObjects 1 "ComObjs at Start" $printChecks

    # Generate the figures for the user manual from the PowerPoint file.
    if { $optExportPpt } {
        puts "    Exporting figures from PowerPoint template ..."
        ::Ppt::ExportPptFile $pptInFile $outFigureDir "Figure-%02d.png" 1 end "PNG" -1 -1 false false
    }
    Cawt::CheckComObjects 1 "ComObjs after ExportPptFile" $printChecks

    # Copy the user manual template to new location and name.
    # Open the user manual template and perform the following actions:
    #   Fill module specific reference tables.
    #   Insert a table with the test programs available (from folder TestPrograms).
    #   Insert the generated figures replacing the placeholder text.
    # Then save the finished user manual in the Final folder.
    file copy -force $wordInFile $userManFile
    set wordId [::Word::OpenNew]
    set docId [::Word::OpenDocument $wordId $userManFile false]
    ::Word::SetCompatibilityMode $docId $::Word::wdWord2003

    set numTables [::Word::GetNumTables $docId]
    Cawt::CheckComObjects 3 "ComObjs after OpenDocument" $printChecks

    if { $optReplaceRefTables } {
        set moduleList [list "Cawt" "Earth" "Excel" "Explorer" "Matlab" "Ocr" "Outlook" "Ppt" "Word"]
        foreach module $moduleList {
            set procList [lsort -dictionary [info commands ${module}::*]]
            foreach procFullName $procList {
                set procShortName [lindex [split $procFullName ":"] end]
                if { [string match "_*" $procShortName] } {
                    # Internal procedures start with "_". Do not add to table.
                    continue
                }
                lappend procFullNameList($module)  $procFullName
                lappend procShortNameList($module) $procShortName
            }
        }

        for { set n 1 } { $n <= $numTables } {incr n } {
            set tableId [::Word::GetTableIdByIndex $docId $n]
            # Placeholder must be listed in row 2, column 1.
            set cellCont [::Word::GetCellValue $tableId 2 1]
            foreach module $moduleList {
                if { $cellCont eq "%TABLE ${module}%" } {
                    set numProcs [llength $procFullNameList($module)]
                    puts "    Replacing table \"%TABLE ${module}%\" with $numProcs procedure references ..."
                    set numRows [::Word::GetNumRows $tableId]
                    set missingRows [expr {$numProcs - $numRows + 1 }]
                    ::Word::AddRow $tableId end $missingRows

                    set row 2
                    foreach procFullName $procFullNameList($module) procShortName $procShortNameList($module) {
                        set procBody [info body $procFullName]
                        set description "N/A"
                        set strBegin [string first "#" $procBody]
                        if { $strBegin >= 0 } {
                            set strEnd [string first "\n" $procBody [expr { $strBegin + 1 }]]
                            if { $strEnd > $strBegin } {
                                set description [string range $procBody [expr { $strBegin + 1 }] [expr { $strEnd - 1 }]]
                            } else {
                                set description [string range $procBody [expr { $strBegin + 1 }] end]
                            }
                        }
                        set description [string trim $description]
                        ::Word::SetCellValue $tableId $row 1 $procShortName
                        set rangeId [::Word::GetCellRange $tableId $row 1]
                        set url [format "%s#%s" $refUrl $procFullName]
                        ::Word::SetHyperlink $rangeId $url $procShortName
                        ::Word::SetCellValue $tableId $row 2 $description
                        ;;Cawt::Destroy $rangeId
                        incr row
                    }
                }
            }
            ::Cawt::Destroy $tableId
        }
        Cawt::CheckComObjects 3 "ComObjs after ReplaceRefTables" $printChecks
    }

    if { $optReplaceTestPrograms } {
        set placeHolder "%TABLE TestPrograms%"
        set foundPlaceHolder false
        for { set n 1 } { $n <= $numTables } {incr n } {
            if { $foundPlaceHolder } {
                break
            }
            set tableId [::Word::GetTableIdByIndex $docId $n]
            # Placeholder must be listed in row 2, column 1.
            set cellCont [::Word::GetCellValue $tableId 2 1]
            if { $cellCont eq $placeHolder } {
                puts "    Replacing table \"$placeHolder\" with list of test programs ..."
                set testFileList [lsort [glob -directory $testDir Earth* Excel* Explorer* Matlab* Outlook* Ocr* Ppt* Word*]]
                set numRows [::Word::GetNumRows $tableId]
                set missingRows [expr [llength $testFileList] - $numRows + 1]
                ::Word::AddRow $tableId end $missingRows

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
            }
            ::Cawt::Destroy $tableId
        }
        if { ! $foundPlaceHolder } {
            puts "Warning: Placeholder \"$placeHolder\" not available in Word template"
        }
        Cawt::CheckComObjects 3 "ComObjs after ReplaceTestPrograms" $printChecks
    }

    if { $optReplaceFigures } {
        set cropValues(Figure-01)  1.5    ; # Overview CAWT
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
                puts "Warning: Figure $figName not available in Word template"
                continue
            } else {
                puts "    Replacing keyword $placeHolder with figure $figImg ..."
            }
            set imgId [::Word::InsertImage $myRange $fig]
            ::Word::ReplaceString $myRange $placeHolder ""
            if { [info exists cropValues($figName)] } {
                set crop [::Cawt::CentiMetersToPoints $cropValues($figName)]
            } else {
                puts "Warning: No crop value specified for figure $figName"
                set crop 0.0
            }
            ::Word::CropImage $imgId $crop
            ::Cawt::Destroy $imgId
            ::Cawt::Destroy $myRange
        }
        Cawt::CheckComObjects 3 "ComObjs after ReplaceFigures" $printChecks
    }

    if { $optReplaceKeywords } {
        if { ! [::Word::FindString $docId "%VERSION%"] } {
            puts "Warning: %VERSION% not available in Word template"
        }
        ::Word::ReplaceString $docId "%VERSION%" $pkgVersion "all"

        if { ! [::Word::FindString $docId "%YEAR%"] } {
            puts "Warning: %YEAR% not available in Word template"
        }
        set year [clock format [clock seconds] -format "%Y"]
        ::Word::ReplaceString $docId "%YEAR%" $year "all"

        if { ! [::Word::FindString $docId "%DATE%"] } {
            puts "Warning: %DATE% not available in Word template"
        }
        set date [clock format [clock seconds] -format "%Y-%m-%d"]
        ::Word::ReplaceString $docId "%DATE%" $date "all"

        Cawt::CheckComObjects 3 "ComObjs after ReplaceKeywords" $printChecks
    }

    ::Word::UpdateFields $docId

    Cawt::CheckComObjects 3 "ComObjs after UpdateFields" $printChecks

    ::Word::SaveAs $docId $userManFile
    set retVal [catch {::Word::SaveAsPdf $docId $pdfManFile} errMsg]
    if { $retVal } {
        puts "Warning: $errMsg"
    }
    ::Word::Close $docId
    ::Word::Quit $wordId
    ::Cawt::Destroy
}
puts "Done."
exit 0
