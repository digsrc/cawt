# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval ::Excel {

    proc Search { worksheetId str { row1 1 } { col1 1 } { row2 -1 } { col2 -1 } } {
        # Find a string in a worksheet cell range.
        #
        # worksheetId - Identifier of the worksheet.
        # str         - Search string.
        # row1        - Row number of upper-left corner of the cell range.
        # col1        - Column number of upper-left corner of the cell range.
        # row2        - Row number of lower-right corner of the cell range.
        # col2        - Column number of lower-right corner of the cell range.
        #
        # If row2 or col2 is negative, all used rows and columns are searched.
        #
        # Return the first matching cell as a 2-element list {row, col} of indices.
        # If no cell matches the search criteria, an empty list is returned.

        if { $row2 < 0 } {
            set row2 [::Excel::GetLastUsedRow $worksheetId]
        }
        if { $col2 < 0 } {
            set col2 [::Excel::GetLastUsedColumn $worksheetId]
        }

        set matrixList [::Excel::GetMatrixValues $worksheetId $row1 $col1 $row2 $col2]
        set row 1
        foreach rowList $matrixList {
            set col [lsearch -exact $rowList $str]
            if { $col >= 0 } {
                return [list $row [expr {$col + 1}]]
            }
            incr row
        }
        return [list]
    }

    proc _ConvertFormula { workbookId formula } {
        set worksheetId [::Excel::GetWorksheetIdByIndex $workbookId 1]
        set cell [::Excel::SelectCellByIndex $worksheetId 1 1]
        set original [$cell Formula]
        $cell Formula $formula
        set formula  [$cell FormulaLocal]
        $cell Formula $original;
        return $formula
    }

    # Generate Absolute Reference Formula of Worksheet.
    proc _ToAbsoluteReference { worksheetId } {
        return [format "\[%s\]%s" [$worksheetId -with { Parent } Name] \
                                  [::Excel::GetWorksheetName $worksheetId]]
    }

    proc DiffExcelFiles { excelBaseFile excelNewFile { r 255 } { g 0 } { b 0 } } {
        # Compare 2 Excel files visually.
        #
        # excelBaseFile - Name of the base Excel file.
        # excelNewFile  - Name of the new Excel file.
        # r             - Red component of the mark color.
        # g             - Green component of the mark color.
        # b             - Blue component of the mark color.
        #
        # The two Excel files are opened in read-only mode and the cells,
        # which are different in the two Excel files are set to the mark color.
        # As Excel does not allow to load two files with identical names (even
        # from different directories), in that case the excelBaseFile is copied
        # into a temporary directory and renamed.
        #
        # The algorithm used is identical to the Excel diff script used in Tortoise SVN.
        #
        # The r, g and b values are specified as integers in the
        # range [0, 255].
        #
        # Return the identifier of the new Excel application instance.
        #
        # See also: OpenNew

        set fastMode false

        if { ! [file exists $excelBaseFile] } {
            error "Diff: Base file $excelBaseFile does not exists"
        }
        if { ! [file exists $excelNewFile] } {
            error "Diff: New file $excelNewFile does not exists"
        }
        if { [file normalize $excelBaseFile] eq [file normalize $excelNewFile] } {
            error "Diff: Base and new file are equal. Cannot compare."
        }
        if { [file tail $excelBaseFile] eq [file tail $excelNewFile] } {
            set tailName  [file tail $excelBaseFile]
            set rootName  [file rootname $tailName]
            set extension [file extension $tailName]
            set tmpName   [format "%s_BaseTmp%s" $rootName $extension]
            set tmpName   [file join [poMisc GetTmpDir] $tmpName]
            file copy -force $excelBaseFile $tmpName
            set excelBaseFile $tmpName
        }
        set appId [::Excel::OpenNew true]
        set baseWorkbookId [::Excel::OpenWorkbook $appId [file nativename $excelBaseFile] true]
        set newWorkbookId  [::Excel::OpenWorkbook $appId [file nativename $excelNewFile]  true]
        if { [::Excel::IsWorkbookProtected $baseWorkbookId] } {
            puts "Unable to arrange windows, because $excelBaseFile is protected."
        }
        if { [::Excel::IsWorkbookProtected $newWorkbookId] } {
            puts "Unable to arrange windows, because $excelNewFile is protected."
        }

        set winId [$appId -with { Windows } Item [expr 2]]
        set caption [$winId Caption]
        $appId -with { Windows } CompareSideBySideWith $caption
        ::Excel::SetWindowState $appId $::Excel::xlMaximized
        ::Excel::ArrangeWindows $appId $::Excel::xlArrangeStyleHorizontal
        if { ! $fastMode && [::Excel::IsWorkbookProtected $newWorkbookId] } {
            puts "Fall back to fast mode because [::Excel::GetWorkbookName $newWorkbookId] is protected.");
            set fastMode true
        }

        # Create a special workbook for formula convertion.
        set convWorkbookId [::Excel::AddWorkbook $appId]

        set numWorksheets [::Excel::GetNumWorksheets $newWorkbookId]
        for { set i 1 } { $i <= $numWorksheets } { incr i } {
            set baseWorksheetId [::Excel::GetWorksheetIdByIndex $baseWorkbookId $i]
            set newWorksheetId  [::Excel::GetWorksheetIdByIndex $newWorkbookId  $i]

            ::Excel::UnhideWorksheet $baseWorksheetId
            ::Excel::UnhideWorksheet $newWorksheetId

            if { ! $fastMode } {
                set lastWorksheetId [::Excel::GetWorksheetIdByIndex $newWorkbookId [::Excel::GetNumWorksheets $newWorkbookId]]
                set dummyWorksheetId [::Excel::CopyWorksheetAfter $baseWorksheetId $lastWorksheetId "Dummy_for_Comparison_$i"]
                $dummyWorksheetId Visible [expr $::Excel::xlSheetVisible]
                ::Excel::SetWorksheetTabColor $dummyWorksheetId 127 127 255
            }

            if { [::Excel::IsWorksheetProtected $newWorksheetId] } {
                puts "Unable to mark differences because the Worksheet is protected."
            } else {
                $newWorksheetId -with { Cells FormatConditions } Delete

                if { $fastMode } {
                    set formula [format "=INDIRECT(\"%s!\"&ADDRESS(ROW(),COLUMN()))" [::Excel::_ToAbsoluteReference $baseWorksheetId]]
                } else {
                    set formula [format "=INDIRECT(\"Dummy_for_Comparison_%d!\"&ADDRESS(ROW(),COLUMN()))" $i]
                }
                set formula [::Excel::_ConvertFormula $convWorkbookId $formula]
                $newWorksheetId -with { Cells FormatConditions } Add $::Excel::xlCellValue $::Excel::xlNotEqual $formula
                set formatCondition [$newWorksheetId -with { Cells FormatConditions } Item 1]
                $formatCondition -with { Interior } Color [::Cawt::RgbToColor $r $g $b]
            }
        }

        # Close the special workbook quietly
        $convWorkbookId Saved [::Cawt::TclBool true]
        $convWorkbookId Close

        # Activate first Worksheet
        ::Excel::GetWorksheetIdByIndex $baseWorkbookId 1 true
        ::Excel::GetWorksheetIdByIndex $newWorkbookId  1 true

        # Suppress save dialog if nothing changed
        $baseWorkbookId Saved [::Cawt::TclBool true]
        $newWorkbookId  Saved [::Cawt::TclBool true]

        return $appId
    }

    proc SetHeaderRow { worksheetId headerList { row 1 } { startCol 1 } } {
        # Insert row values from a Tcl list and format as a header row.
        #
        # worksheetId - Identifier of the worksheet.
        # headerList  - List of values to be inserted as header.
        # row         - Row number. Row numbering starts with 1.
        # startCol    - Column number of insertion start. Column numbering starts with 1.
        #
        # No return value. If headerList is an empty list, an error is thrown.
        #
        # See also: SetRowValues FormatHeaderRow

        set len [llength $headerList]
        ::Excel::SetRowValues $worksheetId $row $headerList $startCol $len
        ::Excel::FormatHeaderRow $worksheetId $row $startCol [expr {$startCol + $len -1}]
    }

    proc FormatHeaderRow { worksheetId row startCol endCol } {
        # Format a row as a header row.
        #
        # worksheetId - Identifier of the worksheet.
        # row         - Row number. Row numbering starts with 1.
        # startCol    - Column number of formatting start. Column numbering starts with 1.
        # endCol      - Column number of formatting end. Column numbering starts with 1.
        #
        # The cell values of a header are formatted as bold text with both vertical and
        # horizontal centered alignment.
        #
        # No return value.
        #
        # See also: SetHeaderRow

        set header [::Excel::SelectRangeByIndex $worksheetId $row $startCol $row $endCol]
        ::Excel::SetRangeHorizontalAlignment $header $::Excel::xlCenter
        ::Excel::SetRangeVerticalAlignment   $header $::Excel::xlCenter
        ::Excel::SetRangeFontBold $header
        ::Cawt::Destroy $header
    }

    proc ClipboardToMatrix { { sepChar ";" } } {
        # Return the matrix values contained in the clipboard.
        #
        # sepChar - The separation character of the clipboard matrix data.
        #
        # The clipboard data must be in CSV format with sepChar as separation character.
        # See SetMatrixValues for the description of a matrix representation.
        #
        # See also: ClipboardToWorksheet MatrixToClipboard

        set csvFmt [twapi::register_clipboard_format "Csv"]
        while { ! [twapi::clipboard_format_available $csvFmt] } {
            after 10
        }
        twapi::open_clipboard
        set clipboardData [twapi::read_clipboard $csvFmt]
        twapi::close_clipboard

        ::Excel::SetCsvSeparatorChar $sepChar
        set matrixList [::Excel::CsvStringToMatrix $clipboardData]
        return $matrixList
    }

    proc ClipboardToWorksheet { worksheetId { startRow 1 } { startCol 1 } { sepChar ";" } } {
        # Insert the matrix values contained in the clipboard into a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        # startRow    - Row number. Row numbering starts with 1.
        # startCol    - Column number. Column numbering starts with 1.
        # sepChar     - The separation character of the clipboard matrix data.
        #
        # The insertion of the matrix data starts at cell "startRow,startCol".
        # The clipboard data must be in CSV format with sepChar as separation character.
        # See SetMatrixValues for the description of a matrix representation.
        #
        # No return value.
        #
        # See also: ClipboardToMatrix WorksheetToClipboard

        set matrixList [::Excel::ClipboardToMatrix $sepChar]
        SetMatrixValues $worksheetId $matrixList $startRow $startCol
    }

    proc MatrixToClipboard { matrixList { sepChar ";" } } {
        # Copy a matrix into the clipboard.
        #
        # matrixList - Matrix with table data.
        # sepChar    - The separation character of the clipboard matrix data.
        #
        # The clipboard data will be in CSV format with sepChar as separation character.
        # See SetMatrixValues for the description of a matrix representation.
        #
        # No return value.
        #
        # See also: WorksheetToClipboard ClipboardToMatrix

        set csvFmt [twapi::register_clipboard_format "Csv"]
        twapi::open_clipboard
        twapi::empty_clipboard
        ::Excel::SetCsvSeparatorChar $sepChar
        twapi::write_clipboard $csvFmt [::Excel::MatrixToCsvString $matrixList]
        twapi::close_clipboard
    }

    proc WorksheetToClipboard { worksheetId row1 col1 row2 col2 { sepChar ";" } } {
        # Copy worksheet data into the clipboard.
        #
        # worksheetId - Identifier of the worksheet.
        # row1        - Row number of upper-left corner of the copy range.
        # col1        - Column number of upper-left corner of the copy range.
        # row2        - Row number of lower-right corner of the copy range.
        # col2        - Column number of lower-right corner of the copy range.
        # sepChar     - The separation character of the clipboard matrix data.
        #
        # The clipboard data will be in CSV format with sepChar as separation character.
        #
        # No return value.
        #
        # See also: ClipboardToWorksheet MatrixToClipboard

        set matrixList [::Excel::GetMatrixValues $worksheetId $row1 $col1 $row2 $col2]
        ::Excel::MatrixToClipboard $matrixList $sepChar
    }
}
