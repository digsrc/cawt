# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval ::Excel {

    proc _MediaWikiList2RowString { lineList sep } {
        if { $sep eq "||" } {
            set lineStr "|-\n| "
        } else {
            set lineStr "|-\n! "
        }
        set len [expr {[llength $lineList] -1}]
        set curVal 0
        foreach val $lineList {
            set tmp [string map {"\n\r" "<br/>"} $val]
            set tmp [string map {"|" "<nowiki>|</nowiki>"} $tmp]
            append lineStr $tmp
            if { $curVal < $len } {
                append lineStr " $sep "
            }
            incr curVal
        }
        return $lineStr
    }

    proc _MediaWikiSubstHtml { word } {
        set tmp [string trim $word]
        # Substitute <nowiki> and <br/> keywords.
        set tmp [string map {"<br/>" "\n\r" } $tmp]
        set tmp [string map {"<nowiki>|</nowiki>" "|" } $tmp]
        return $tmp
    }

    proc _MediaWikiRowString2List { line sep } {
        set tmpList {}
        set range [string range $line 1 end]
        while { [string first $sep $range] >= 0 } {
            set begRange 0
            set endRange [expr {[string first $sep $range] - 1}]
            lappend tmpList [::Excel::_MediaWikiSubstHtml [string range $range $begRange $endRange]]
            # Set new range to start after the separator. We add 3, because the endRange
            # index points to the character before the separator.
            set range [string range $range [expr {$endRange+3}] end]
        }
        lappend tmpList [::Excel::_MediaWikiSubstHtml [string range $range 0 end]]
        return $tmpList
    }

    proc ReadMediaWikiFile { wikiFileName { useHeader true } } {
        # Read a MediaWiki table file into a matrix.
        #
        # wikiFileName - Name of the MediaWiki file.
        # useHeader    - true: Insert the header of the MediaWiki table as first row.
        #                false: Only transfer the table data.
        #
        # Return the MediaWiki table data as a matrix.
        # See SetMatrixValues for the description of a matrix representation.
        #
        # See also: WriteMediaWikiFile MediaWikiFileToWorksheet

        set tmpList  {}
        set matrixList {}
        set firstRow true

        set catchVal [catch {open $wikiFileName r} fp]
        if { $catchVal != 0 } {
            error "Could not open file \"$wikiFileName\" for reading."
        }

        while { [gets $fp line] >= 0 } {
            if { [string index $line 0] eq "!" && $useHeader } {
                set tmpList [::Excel::_MediaWikiRowString2List $line "!!"]
            } elseif { [string range $line 0 1] eq "|-" || \
                       [string range $line 0 1] eq "|\}" } {
                if { $firstRow } {
                    set firstRow false
                    continue
                }
                lappend matrixList $tmpList
                set tmpList {}
            } elseif { [string index $line 0] eq "|" } {
                if { [string first "||" $line] >= 0 } {
                    set tmpList [::Excel::_MediaWikiRowString2List $line "||"]
                } else {
                    lappend tmpList [::Excel::_MediaWikiSubstHtml [string range $line 1 end]]
                }
            }
        }
        close $fp
        return $matrixList
    }

    proc _WriteMediaWikiHeader { fp headerList { tableName "" } } {
        puts $fp "\{| class=\"wikitable border=\"1\""
        if { $tableName ne "" } {
            puts $fp "|+ $tableName"
        }
        puts $fp [::Excel::_MediaWikiList2RowString $headerList "!!"]
    }

    proc _WriteMediaWikiData { fp matrixList } {
        foreach row $matrixList {
            puts $fp [::Excel::_MediaWikiList2RowString $row "||"]
        }
        puts $fp "|\}"
    }

    proc WriteMediaWikiFile { matrixList wikiFileName { useHeader true } { tableName "" } } {
        # Write the values of a matrix into a MediaWiki table file.
        #
        # matrixList    - Matrix with table data.
        # wikiFileName  - Name of the MediaWiki file.
        # useHeader     - true: Use first row of the matrix as header of the
        #                 MediaWiki table.
        #                 false: Only transfer the table data.
        # tableName     - Table name (caption) of the generated MediaWiki table.
        #
        # See SetMatrixValues for the description of a matrix representation.
        #
        # No return value.
        #
        # See also: ReadMediaWikiFile WorksheetToMediaWikiFile

        set catchVal [catch {open $wikiFileName w} fp]
        if { $catchVal != 0 } {
            error "Could not open file \"$wikiFileName\" for writing."
        }

        puts $fp "{| class=\"wikitable border=\"1\""
        if { $tableName ne "" } {
            puts $fp "|+ $tableName"
        }
        set curLine 1
        foreach line $matrixList {
            if { $useHeader && $curLine == 1 } {
                puts $fp [::Excel::_MediaWikiList2RowString $line "!!"]
            } else {
                puts $fp [::Excel::_MediaWikiList2RowString $line "||"]
            }
            incr curLine
        }
        puts $fp "|}"
        close $fp
    }

    proc MediaWikiFileToWorksheet { wikiFileName worksheetId { useHeader true } } {
        # Insert the values of a MediaWiki table file into a worksheet.
        #
        # wikiFileName - Name of the MediaWiki file.
        # worksheetId  - Identifier of the worksheet.
        # useHeader    - true: Insert the header of the MediaWiki table as first row.
        #                false: Only transfer the table data.
        #
        # The insertion starts at row and column 1.
        # Values contained in the worksheet cells are overwritten.
        #
        # No return value.
        #
        # See also: WorksheetToMediaWikiFile SetMatrixValues
        # WikitFileToWorksheet WordTableToWorksheet MatlabFileToWorksheet
        # RawImageFileToWorksheet TablelistToWorksheet

        set catchVal [catch {open $wikiFileName "r"} fp]
        if { $catchVal != 0 } {
            error "Could not open file \"$wikiFileName\" for reading."
        }

        # TODO |- style="background:green"
        set row      1
        set firstRow true
        set rowList  {}

        while { [gets $fp line] >= 0 } {
            if { [string index $line 0] eq "!" && $useHeader } {
                # Found a header line. Currently only headers with "!!" separators are supported.
                set headerList [::Excel::_MediaWikiRowString2List $line "!!"]
                ::Excel::SetHeaderRow $worksheetId $headerList
                incr row
            } elseif { [string range $line 0 1] eq "|+" } {
                set worksheetName [string trim [string range $line 2 end]]
                Excel::SetWorksheetName $worksheetId $worksheetName
            } elseif { [string range $line 0 1] eq "|-" || \
                       [string range $line 0 1] eq "|\}" } {
                if { $firstRow } {
                    set firstRow false
                    continue
                }
                if { [llength $rowList] != 0 } {
                    ::Excel::SetRowValues $worksheetId $row $rowList
                    incr row
                }
                set rowList {}
            } elseif { [string index $line 0] eq "|" } {
                if { [string first "||" $line] >= 0 } {
                    set rowList [::Excel::_MediaWikiRowString2List $line "||"]
                } else {
                    lappend rowList [::Excel::_MediaWikiSubstHtml [string range $line 1 end]]
                }
            }
        }
        close $fp
    }

    proc WorksheetToMediaWikiFile { worksheetId wikiFileName { useHeader true } } {
        # Insert the values of a worksheet into a MediaWiki table file.
        #
        # worksheetId  - Identifier of the worksheet.
        # wikiFileName - Name of the MediaWiki file.
        # useHeader    - true:  Use the first row of the worksheet as the header
        #                       of the MediaWiki table.
        #                false: Do not generate a MediaWiki table header. All worksheet
        #                       cells are interpreted as data.
        #
        # No return value.
        #
        # See also: MediaWikiFileToWorksheet GetMatrixValues
        # WorksheetToWikitFile WorksheetToWordTable WorksheetToMatlabFile
        # WorksheetToRawImageFile WorksheetToTablelist

        set numRows [::Excel::GetLastUsedRow $worksheetId]
        set numCols [::Excel::GetLastUsedColumn $worksheetId]
        set startRow 1
        set catchVal [catch {open $wikiFileName w} fp]
        if { $catchVal != 0 } {
            error "Could not open file \"$wikiFileName\" for writing."
        }
        if { $useHeader } {
            set headerList [::Excel::GetMatrixValues $worksheetId $startRow 1 $startRow $numCols]
            set worksheetName [::Excel::GetWorksheetName $worksheetId]
            ::Excel::_WriteMediaWikiHeader $fp [lindex $headerList 0] $worksheetName
            incr startRow
        }
        set matrixList [::Excel::GetMatrixValues $worksheetId $startRow 1 $numRows $numCols]
        ::Excel::_WriteMediaWikiData $fp $matrixList
        close $fp
    }

    proc MediaWikiFileToExcelFile { wikiFileName excelFileName \
                                   { useHeader true } { quitExcel true } } {
        # Convert a MediaWiki table file to an Excel file.
        #
        # wikiFileName  - Name of the MediaWiki input file.
        # excelFileName - Name of the Excel output file.
        # useHeader     - true:  Use header information from the MediaWiki file to
        #                        generate an Excel header (see SetHeaderRow).
        #                 false: Only transfer the table data.
        # quitExcel     - true:  Quit the Excel instance after generation of output file.
        #                 false: Leave the Excel instance open after generation of output file.
        #
        # The table data from the MediaWiki file will be inserted into a worksheet named "MediaWiki".
        #
        # Return the Excel application identifier, if quitExcel is false.
        # Otherwise no return value.
        #
        # See also: MediaWikiFileToWorksheet ExcelFileToMediaWikiFile
        # ReadMediaWikiFile WriteMediaWikiFile WikitFileToExcelFile

        set appId [::Excel::OpenNew true]
        set workbookId [::Excel::AddWorkbook $appId]
        set worksheetId [::Excel::AddWorksheet $workbookId "MediaWiki"]
        ::Excel::MediaWikiFileToWorksheet $wikiFileName $worksheetId $useHeader
        ::Excel::SaveAs $workbookId $excelFileName
        if { $quitExcel } {
            ::Excel::Quit $appId
        } else {
            return $appId
        }
    }

    proc ExcelFileToMediaWikiFile { excelFileName wikiFileName { worksheetNameOrIndex 0 } \
                                   { useHeader true } { quitExcel true } } {
        # Convert an Excel file to a MediaWiki table file.
        #
        # excelFileName        - Name of the Excel input file.
        # wikiFileName         - Name of the MediaWiki output file.
        # worksheetNameOrIndex - Worksheet name or index to convert.
        # useHeader            - true:  Use the first row of the worksheet as the header
        #                               of the MediaWiki table.
        #                        false: Do not generate a MediaWiki table header. All worksheet
        #                               cells are interpreted as data.
        # quitExcel            - true:  Quit the Excel instance after generation of output file.
        #                        false: Leave the Excel instance open after generation of output file.
        #
        # Note, that the Excel Workbook is opened in read-only mode.
        #
        # Return the Excel application identifier, if quitExcel is false.
        # Otherwise no return value.
        #
        # See also: MediaWikiFileToWorksheet MediaWikiFileToExcelFile
        # ReadMediaWikiFile WriteMediaWikiFile WikitFileToExcelFile

        set appId [::Excel::OpenNew true]
        set workbookId [::Excel::OpenWorkbook $appId $excelFileName true]
        if { [string is integer $worksheetNameOrIndex] } {
            set worksheetId [::Excel::GetWorksheetIdByIndex $workbookId [expr int($worksheetNameOrIndex)]]
        } else {
            set worksheetId [::Excel::GetWorksheetIdByName $workbookId $worksheetNameOrIndex]
        }
        ::Excel::WorksheetToMediaWikiFile $worksheetId $wikiFileName $useHeader
        if { $quitExcel } {
            ::Excel::Quit $appId
        } else {
            return $appId
        }
    }
}
