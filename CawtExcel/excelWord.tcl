# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval Excel {

    namespace ensemble create

    namespace export WordTableToWorksheet
    namespace export WorksheetToWordTable

    proc WordTableToWorksheet { tableId worksheetId { useHeader true } } {
        # Insert the values of a Word table into a worksheet.
        #
        # tableId     - Identifier of the Word table.
        # worksheetId - Identifier of the worksheet.
        # useHeader   - true: Insert the header of the Word table as first row.
        #               false: Only transfer the table data.
        #
        # No return value.
        #
        # See also: WorksheetToWordTable SetMatrixValues
        # WikitFileToWorksheet MediaWikiFileToWorksheet MatlabFileToWorksheet
        # RawImageFileToWorksheet TablelistToWorksheet

        set numCols [Word GetNumColumns $tableId]
        if { $useHeader } {
            for { set col 1 } { $col <= $numCols } { incr col } {
                lappend headerList [Word GetCellValue $tableId 1 $col]
            }
            Excel SetHeaderRow $worksheetId $headerList
        }
        set numRows [Word GetNumRows $tableId]
        incr numRows -1
        set startWordRow 2
        if { $useHeader } {
            set startExcelRow 2
        } else {
            set startExcelRow 1
        }
        set tableList [Word GetMatrixValues $tableId \
                      $startWordRow 1 [expr {$startWordRow + $numRows-1}] $numCols]
        Excel SetMatrixValues $worksheetId $tableList $startExcelRow 1
    }

    proc WorksheetToWordTable { worksheetId tableId { useHeader true } } {
        # Insert the values of a worksheet into a Word table.
        #
        # worksheetId - Identifier of the worksheet.
        # tableId     - Identifier of the Word table.
        # useHeader   - true: Use the first row of the worksheet as the header
        #                     of the Word table.
        #               false: Do not generate a Word table header. All worksheet
        #                      cells are interpreted as data.
        #
        # No return value.
        #
        # See also: WordTableToWorksheet GetMatrixValues
        # WorksheetToWikitFile WorksheetToMediaWikiFile WorksheetToMatlabFile
        # WorksheetToRawImageFile WorksheetToTablelist

        set numRows [Excel GetLastUsedRow $worksheetId]
        set numCols [Excel GetLastUsedColumn $worksheetId]
        set startRow 1
        set headerList [Excel GetRowValues $worksheetId 1 1 $numCols]
        if { [llength $headerList] < $numCols } {
            set numCols [llength $headerList]
        }
        if { $useHeader } {
            Word SetHeaderRow $tableId $headerList
            incr startRow
        }
        set excelList [Excel GetMatrixValues $worksheetId $startRow 1 $numRows $numCols]
        Word SetMatrixValues $tableId $excelList $startRow 1
    }
}
