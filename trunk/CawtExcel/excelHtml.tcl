# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval Excel {

    namespace ensemble create

    namespace export ExcelFileToHtmlFile
    namespace export WorksheetToHtmlFile
    namespace export WriteHtmlFile

    proc _WriteTableBegin { fp } {
        puts $fp "<table>"
    }

    proc _WriteTableEnd { fp } {
        puts $fp "</table>"
    }

    proc _WriteRowBegin { fp } {
        puts $fp "  <tr>"
    }

    proc _WriteRowEnd { fp } {
        puts $fp "  </tr>"
    }

    proc _WriteTableCell { fp tag val { bgColor "" } { fgColor "" } } {
        # TODO Write fgColor
        puts $fp "    <$tag bgcolor=\"$bgColor\">$val</$tag>"
    }

    proc _WriteTableRow { fp rowValues { tag "td" } { bgColor "" } { fgColor "" } } {
        Excel::_WriteRowBegin $fp
        foreach val $rowValues {
            Excel::_WriteTableCell $fp $tag $val $bgColor $fgColor
        }
        Excel::_WriteRowEnd $fp
    }

    proc _WriteTableHeader { fp rowValues } {
        Excel::_WriteTableRow $fp $rowValues "th"
    }

    proc _WriteTableData { fp matrixValues } {
        foreach row $matrixValues {
            Excel::_WriteTableRow $fp $row
        }
    }

    proc _WriteCell { fp tag worksheetId row col } {
        set val [Excel GetCellValue $worksheetId $row $col]
        set cellId [Excel GetCellIdByIndex $worksheetId $row $col]
        set bgRgb  [Excel GetRangeFillColor $cellId]
        set fgRgb  [Excel GetRangeTextColor $cellId]
        set bgHex  [format "#%02X%02X%02X" [lindex $bgRgb 0] [lindex $bgRgb 1] [lindex $bgRgb 2]]
        set fgHex  [format "#%02X%02X%02X" [lindex $fgRgb 0] [lindex $fgRgb 1] [lindex $fgRgb 2]]
        Excel::_WriteTableCell $fp $tag $val $bgHex $fgHex
    }

    proc WriteHtmlFile { matrixList htmlFileName { useHeader true } } {
        # Write the values of a matrix into a Html table file.
        #
        # matrixList    - Matrix with table data.
        # htmlFileName  - Name of the HTML file.
        # useHeader     - true: Use first row of the matrix as header of the
        #                 HTML table.
        #
        # See SetMatrixValues for the description of a matrix representation.
        #
        # No return value.
        #
        # See also: WorksheetToHtmlFile

        set catchVal [catch {open $htmlFileName w} fp]
        if { $catchVal != 0 } {
            error "Could not open file \"$htmlFileName\" for writing."
        }

        Excel::_WriteTableBegin $fp
        set curRow 1
        foreach row $matrixList {
            if { $useHeader && $curRow == 1 } {
                Excel::_WriteTableHeader $fp $row
            } else {
                Excel::_WriteTableRow $fp $row
            }
            incr curRow
        }
        Excel::_WriteTableEnd $fp
        close $fp
    }

    proc WorksheetToHtmlFile { worksheetId htmlFileName { useHeader true } } {
        # Insert the values of a worksheet into a HTML table file.
        #
        # worksheetId  - Identifier of the worksheet.
        # htmlFileName - Name of the HTML file.
        # useHeader    - true:  Use the first row of the worksheet as the header
        #                       of the HTML table.
        #                false: Do not generate a HTML table header. All worksheet
        #                       cells are interpreted as data.
        #
        # No return value.
        #
        # See also: GetMatrixValues
        # WorksheetToMediaWikiFile WorksheetToWikitFile WorksheetToWordTable
        # WorksheetToMatlabFile WorksheetToRawImageFile WorksheetToTablelist

        set numRows [Excel GetLastUsedRow $worksheetId]
        set numCols [Excel GetLastUsedColumn $worksheetId]
        set startRow 1
        set catchVal [catch {open $htmlFileName w} fp]
        if { $catchVal != 0 } {
            error "Could not open file \"$htmlFileName\" for writing."
        }
        Excel::_WriteTableBegin $fp
        if { $useHeader } {
            Excel::_WriteRowBegin $fp
            for { set col 1 } { $col <= $numCols } { incr col } {
                Excel::_WriteCell $fp "th" $worksheetId 1 $col
            }
            Excel::_WriteRowEnd $fp
            incr startRow
        }
        for { set row $startRow } { $row <= $numRows } { incr row } {
            Excel::_WriteRowBegin $fp
            for { set col 1 } { $col <= $numCols } { incr col } {
                Excel::_WriteCell $fp "td" $worksheetId $row $col
            }
            Excel::_WriteRowEnd $fp
        }
        Excel::_WriteTableEnd $fp
        close $fp
    }

    proc ExcelFileToHtmlFile { excelFileName htmlFileName { worksheetNameOrIndex 0 } \
                               { useHeader true } { quitExcel true } } {
        # Convert an Excel file to a HTML table file.
        #
        # excelFileName        - Name of the Excel input file.
        # htmlFileName         - Name of the HTML output file.
        # worksheetNameOrIndex - Worksheet name or index to convert.
        # useHeader            - true:  Use the first row of the worksheet as the header
        #                               of the HTML table.
        #                        false: Do not generate a HTML table header. All worksheet
        #                               cells are interpreted as data.
        # quitExcel            - true:  Quit the Excel instance after generation of output file.
        #                        false: Leave the Excel instance open after generation of output file.
        #
        # Note, that the Excel Workbook is opened in read-only mode.
        #
        # Return the Excel application identifier, if quitExcel is false.
        # Otherwise no return value.
        #
        # See also: WriteHtmlFile MediaWikiFileToExcelFile WikitFileToExcelFile

        set appId [Excel OpenNew true]
        set workbookId [Excel OpenWorkbook $appId $excelFileName true]
        if { [string is integer $worksheetNameOrIndex] } {
            set worksheetId [Excel GetWorksheetIdByIndex $workbookId [expr int($worksheetNameOrIndex)]]
        } else {
            set worksheetId [Excel GetWorksheetIdByName $workbookId $worksheetNameOrIndex]
        }
        Excel WorksheetToHtmlFile $worksheetId $htmlFileName $useHeader
        if { $quitExcel } {
            Excel Quit $appId
        } else {
            return $appId
        }
    }
}
