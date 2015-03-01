# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

# Format of Wikit tables as used in the Tcl'ers Wiki:
#
# %| Name           | Home page           | Nick name | Remarks |%   Header row
# &|[Paul Obermeier]|http://www.poSoft.de/|[PO]       |None     |&   Row with odd/even coloring
#  |[Paul Obermeier]|http://www.poSoft.de/|[PO]       |None     |    Standard row
#
# Keywords: <<pipe>> <<br>>

namespace eval Excel {

    namespace ensemble create

    namespace export ExcelFileToWikitFile
    namespace export ReadWikitFile
    namespace export WikitFileToExcelFile
    namespace export WikitFileToWorksheet
    namespace export WorksheetToWikitFile
    namespace export WriteWikitFile

    proc _WikitList2RowString { lineList beginSep endSep } {
        set lineStr "$beginSep "
        set len [expr {[llength $lineList] -1}]
        set curVal 0
        foreach val $lineList {
            set tmp [string map {\n\r "<<br>>"} $val]
            set tmp [string map {"|" "<<pipe>>"} $tmp]
            append lineStr $tmp
            if { $curVal < $len } {
                append lineStr " | "
            }
            incr curVal
        }
        append lineStr $endSep
        return $lineStr
    }

    proc _WikitSubstHtml { word } {
        set tmp [string trim $word]
        set tmp [string map {"<<br>>" "\n\r" } $tmp]
        set tmp [string map {"<<pipe>>" "|" } $tmp]
        return $tmp
    }

    proc _WikitRowString2List { rowStr } {
        set rowList [list]
        foreach cell [split $rowStr "|"] {
            lappend rowList [Excel::_WikitSubstHtml $cell]
        }
        return $rowList
    }

    proc ReadWikitFile { wikiFileName { useHeader true } } {
        # Read a Wikit table file into a matrix.
        #
        # wikiFileName - Name of the Wikit file.
        # useHeader    - true: Insert the header of the Wikit table as first row.
        #                false: Only transfer the table data.
        #
        # Return the Wikit table data as a matrix.
        # See SetMatrixValues for the description of a matrix representation.
        #
        # See also: WriteWikitFile WikitFileToWorksheet

        set catchVal [catch {open $wikiFileName r} fp]
        if { $catchVal != 0 } {
            error "Could not open file \"$wikiFileName\" for reading."
        }

        set matrixList [list]
        while { [gets $fp line] >= 0 } {
            set line [string trim $line]
            if { ( [string range $line 0 1] eq "%|" && $useHeader ) } {
                set rowStr [string map {"%|" "" "|%" "" } $line]
            } elseif { [string range $line 0 1] eq "&|" } {
                set rowStr [string map {"&|" "" "|&" "" } $line]
            } elseif { [string index $line 0] eq "|" } {
                set rowStr [string range $line 1 end-1]
            }
            lappend matrixList [Excel::_WikitRowString2List $rowStr]
        }
        close $fp
        return $matrixList
    }

    proc _WriteWikitHeader { fp headerList { tableName "" } } {
        puts $fp [Excel::_WikitList2RowString $headerList "%|" "|%"]
    }

    proc _WriteWikitData { fp matrixList } {
        foreach row $matrixList {
            puts $fp [Excel::_WikitList2RowString $row "&|" "|&"]
        }
    }

    proc WriteWikitFile { matrixList wikiFileName { useHeader true } } {
        # Write the values of a matrix into a Wikit table file.
        #
        # matrixList    - Matrix with table data.
        # wikiFileName  - Name of the Wikit file.
        # useHeader     - true: Use first row of the matrix as header of the
        #                 Wikit table.
        #
        # See SetMatrixValues for the description of a matrix representation.
        #
        # No return value.
        #
        # See also: ReadWikitFile WorksheetToWikitFile

        set catchVal [catch {open $wikiFileName w} fp]
        if { $catchVal != 0 } {
            error "Could not open file \"$wikiFileName\" for writing."
        }

        set curLine 1
        foreach line $matrixList {
            if { $useHeader && $curLine == 1 } {
                puts $fp [Excel::_WikitList2RowString $line "%|" "|%"]
            } else {
                puts $fp [Excel::_WikitList2RowString $line "&|" "|&"]
            }
            incr curLine
        }
        close $fp
    }

    proc WikitFileToWorksheet { wikiFileName worksheetId { useHeader true } } {
        # Insert the values of a Wikit table file into a worksheet.
        #
        # wikiFileName - Name of the Wikit file.
        # worksheetId  - Identifier of the worksheet.
        # useHeader    - true: Insert the header of the Wikit table as first row.
        #                false: Only transfer the table data.
        #
        # The insertion starts at row and column 1.
        # Values contained in the worksheet cells are overwritten.
        #
        # No return value.
        #
        # See also: WorksheetToWikitFile SetMatrixValues
        # MediaWikiFileToWorksheet WordTableToWorksheet MatlabFileToWorksheet
        # RawImageFileToWorksheet TablelistToWorksheet

        set catchVal [catch {open $wikiFileName "r"} fp]
        if { $catchVal != 0 } {
            error "Could not open file \"$wikiFileName\" for reading."
        }

        set matrixList [list]
        set row 1
        while { [gets $fp line] >= 0 } {
            set line [string trim $line]
            if { ( [string range $line 0 1] eq "%|" && $useHeader ) } {
                set rowStr [string map {"%|" "" "|%" "" } $line]
                set rowList [Excel::_WikitRowString2List $rowStr]
                Excel SetHeaderRow $worksheetId $rowList
            } elseif { [string range $line 0 1] eq "&|" } {
                set rowStr [string map {"&|" "" "|&" "" } $line]
                set rowList [Excel::_WikitRowString2List $rowStr]
                Excel SetRowValues $worksheetId $row $rowList
            } elseif { [string index $line 0] eq "|" } {
                set rowStr [string range $line 1 end-1]
                set rowList [Excel::_WikitRowString2List $rowStr]
                Excel SetRowValues $worksheetId $row $rowList
            }
            incr row
        }
        close $fp
    }

    proc WorksheetToWikitFile { worksheetId wikiFileName { useHeader true } } {
        # Insert the values of a worksheet into a Wikit table file.
        #
        # worksheetId  - Identifier of the worksheet.
        # wikiFileName - Name of the Wikit file.
        # useHeader    - true:  Use the first row of the worksheet as the header
        #                       of the Wikit table.
        #                false: Do not generate a Wikit table header. All worksheet
        #                       cells are interpreted as data.
        #
        # No return value.
        #
        # See also: WikitFileToWorksheet GetMatrixValues
        # WorksheetToMediaWikiFile WorksheetToWordTable WorksheetToMatlabFile
        # WorksheetToRawImageFile WorksheetToTablelist

        set numRows [Excel GetLastUsedRow $worksheetId]
        set numCols [Excel GetLastUsedColumn $worksheetId]
        set startRow 1
        set catchVal [catch {open $wikiFileName w} fp]
        if { $catchVal != 0 } {
            error "Could not open file \"$wikiFileName\" for writing."
        }
        if { $useHeader } {
            set headerList [Excel GetMatrixValues $worksheetId $startRow 1 $startRow $numCols]
            set worksheetName [Excel GetWorksheetName $worksheetId]
            Excel::_WriteWikitHeader $fp [lindex $headerList 0] $worksheetName
            incr startRow
        }
        set matrixList [Excel GetMatrixValues $worksheetId $startRow 1 $numRows $numCols]
        Excel::_WriteWikitData $fp $matrixList
        close $fp
    }

    proc WikitFileToExcelFile { wikiFileName excelFileName \
                                { useHeader true } { quitExcel true } } {
        # Convert a Wikit table file to an Excel file.
        #
        # wikiFileName  - Name of the Wikit input file.
        # excelFileName - Name of the Excel output file.
        # useHeader     - true:  Use header information from the Wikit file to
        #                        generate an Excel header (see SetHeaderRow).
        #                 false: Only transfer the table data.
        # quitExcel     - true:  Quit the Excel instance after generation of output file.
        #                 false: Leave the Excel instance open after generation of output file.
        #
        # The table data from the Wikit file will be inserted into a worksheet named "Wikit".
        #
        # Return the Excel application identifier, if quitExcel is false.
        # Otherwise no return value.
        #
        # See also: WikitFileToWorksheet ExcelFileToWikitFile
        # ReadWikitFile MediaWikiFileToExcelFile

        set appId [Excel OpenNew true]
        set workbookId [Excel AddWorkbook $appId]
        set worksheetId [Excel AddWorksheet $workbookId "Wikit"]
        Excel WikitFileToWorksheet $wikiFileName $worksheetId $useHeader
        Excel SaveAs $workbookId $excelFileName
        if { $quitExcel } {
            Excel Quit $appId
        } else {
            return $appId
        }
    }

    proc ExcelFileToWikitFile { excelFileName wikiFileName { worksheetNameOrIndex 0 } \
                                { useHeader true } { quitExcel true } } {
        # Convert an Excel file to a Wikit table file.
        #
        # excelFileName        - Name of the Excel input file.
        # wikiFileName         - Name of the Wikit output file.
        # worksheetNameOrIndex - Worksheet name or index to convert.
        # useHeader            - true:  Use the first row of the worksheet as the header
        #                               of the Wikit table.
        #                        false: Do not generate a Wikit table header. All worksheet
        #                               cells are interpreted as data.
        # quitExcel            - true:  Quit the Excel instance after generation of output file.
        #                        false: Leave the Excel instance open after generation of output file.
        #
        # Note, that the Excel Workbook is opened in read-only mode.
        #
        # Return the Excel application identifier, if quitExcel is false.
        # Otherwise no return value.
        #
        # See also: WikitFileToWorksheet WikitFileToExcelFile
        # ReadWikitFile WriteWikitFile MediaWikiFileToExcelFile

        set appId [Excel OpenNew true]
        set workbookId [Excel OpenWorkbook $appId $excelFileName true]
        if { [string is integer $worksheetNameOrIndex] } {
            set worksheetId [Excel GetWorksheetIdByIndex $workbookId [expr int($worksheetNameOrIndex)]]
        } else {
            set worksheetId [Excel GetWorksheetIdByName $workbookId $worksheetNameOrIndex]
        }
        Excel WorksheetToWikitFile $worksheetId $wikiFileName $useHeader
        if { $quitExcel } {
            Excel Quit $appId
        } else {
            return $appId
        }
    }
}
