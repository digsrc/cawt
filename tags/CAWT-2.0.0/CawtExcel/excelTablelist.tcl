# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval Excel {

    namespace ensemble create

    namespace export GetTablelistHeader
    namespace export GetTablelistValues
    namespace export SetTablelistHeader
    namespace export SetTablelistValues
    namespace export TablelistToWorksheet
    namespace export WorksheetToTablelist

    proc GetTablelistHeader { tableId } {
        # Return the header line of a tablelist as a list.
        #
        # tableId - Identifier of the tablelist.
        #
        # See also: TablelistToWorksheet WorksheetToTablelist
        # SetTablelistHeader GetTablelistValues

        set numCols [$tableId columncount]
        for { set col 0 } { $col < $numCols } { incr col } {
            lappend headerList [$tableId columncget $col -title]
        }
        return $headerList
    }

    proc GetTablelistValues { tableId } {
        # Return the values of a tablelist as a matrix.
        #
        # tableId - Identifier of the tablelist.
        #
        # See also: TablelistToWorksheet WorksheetToTablelist
        # SetTablelistValues GetTablelistHeader

        return [$tableId get 0 end]
    }

    proc SetTablelistHeader { tableId headerList } {
        # Insert header values into a tablelist.
        #
        # No return value.
        #
        # tableId    - Identifier of the tablelist.
        # headerList - List with table header data.
        #
        # See also: TablelistToWorksheet WorksheetToTablelist
        # SetTablelistValues GetTablelistHeader

        foreach title $headerList {
            $tableId insertcolumns end 0 $title left
        }
    }

    proc SetTablelistValues { tableId matrixList } {
        # Insert matrix values into a tablelist.
        #
        # No return value.
        #
        # tableId    - Identifier of the tablelist.
        # matrixList - Matrix with table data.
        #
        # See also: TablelistToWorksheet WorksheetToTablelist
        # SetTablelistHeader GetTablelistValues

        foreach rowList $matrixList {
            $tableId insert end $rowList
        }
    }

    proc TablelistToWorksheet { tableId worksheetId { useHeader true } { startRow 1 } } {
        # Insert the values of a tablelist into a worksheet.
        #
        # tableId     - Identifier of the tablelist.
        # worksheetId - Identifier of the worksheet.
        # useHeader   - true: Insert the header of the tablelist as first row.
        #               false: Only transfer the tablelist data.
        # startRow    - Row number of insertion start. Row numbering starts with 1.
        #
        # No return value.
        #
        # See also: WorksheetToTablelist SetMatrixValues
        # WikitFileToWorksheet MediaWikiFileToWorksheet MatlabFileToWorksheet
        # RawImageFileToWorksheet WordTableToWorksheet

        set curRow $startRow
        if { $useHeader } {
            set numCols [$tableId columncount]
            for { set col 0 } { $col < $numCols } { incr col } {
                lappend headerList [$tableId columncget $col -title]
            }
            Excel SetHeaderRow $worksheetId $headerList $curRow
            incr curRow
        }
        set matrixList [$tableId get 0 end]
        Excel SetMatrixValues $worksheetId $matrixList $curRow 1
    }

    proc WorksheetToTablelist { worksheetId tableId { useHeader true } } {
        # Insert the values of a worksheet into a tablelist.
        #
        # worksheetId - Identifier of the worksheet.
        # tableId     - Identifier of the tablelist.
        # useHeader   - true: Use the first row of the worksheet as the header
        #                     of the tablelist.
        #               false: Do not generate a tablelist header. All worksheet
        #                      cells are interpreted as data.
        #
        # No return value.
        #
        # See also: TablelistToWorksheet GetMatrixValues
        # WorksheetToWikitFile WorksheetToMediaWikiFile WorksheetToMatlabFile
        # WorksheetToRawImageFile WorksheetToWordTable

        set numRows [Excel GetLastUsedRow $worksheetId]
        set numCols [Excel GetLastUsedColumn $worksheetId]
        set startRow 1
        if { $useHeader } {
            set headerList [Excel GetRowValues $worksheetId 1 1 $numCols]
            foreach title $headerList {
                $tableId insertcolumns end 0 $title left
            }
            incr startRow
        } else {
            for { set col 1 } { $col <= $numCols } { incr col } {
                $tableId insertcolumns end 0 "NN" left
            }
        }
        set excelList [Excel GetMatrixValues $worksheetId $startRow 1 $numRows $numCols]
        foreach rowList $excelList {
            $tableId insert end $rowList
        }
    }
}
