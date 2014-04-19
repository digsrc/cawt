# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval ::Excel {

    proc _PutMatlabHeader { matFp width height matrixName } {
        puts -nonewline $matFp [binary format iiiii \
                                0 $height $width 0 [string length $matrixName]]
        puts -nonewline $matFp [binary format a* $matrixName]
    }

    proc _GetMatlabHeader { matFp } {
        # Check for Level 4 MAT-File.
        set mat4Header [read $matFp 20]
        binary scan $mat4Header iiiii matType height width \
                                      matImaginary matNameLen

        if { $matType == 0    || $matNumRows == 0 || \
             $matNumCols == 0 || $matImaginary == 0 } {
            set matVersion 4
            set matName [read $matFp $matNameLen]
        } else {
            set matVersion 5
            seek $matFp 0
            error "Matlab Level 5 files not yet supported"
        }
        if { $matType != 0 } {
            error "Currently only Intel double-precision numeric matrices are supported."
        }
        return [list $matVersion $width $height]
    }

    proc ReadMatlabHeader { matFileName } {
        # Read the header of a Matlab file.
        #
        # matFileName - Name of the Matlab file.
        #
        # Return the header information as a list of integers containing the
        # following values: MatlabVersion Width Height
        #
        # See also: ReadMatlabFile

        set retVal [catch {open $matFileName "r"} matFp]
        if { $retVal != 0 } {
            error "Cannot open file $matFileName"
        }
        fconfigure $matFp -translation binary
        set headerList [::Excel::_GetMatlabHeader $matFp]
        close $matFp
        return $headerList
    }

    proc ReadMatlabFile { matFileName } {
        # Read a Matlab file into a matrix.
        #
        # matFileName - Name of the Matlab file.
        #
        # Note: Only Matlab Level 4 files are currently supported.
        #
        # Return the Matlab file data as a matrix.
        # See SetMatrixValues for the description of a matrix representation.
        #
        # See also: ReadMatlabHeader WriteMatlabFile MatlabFileToWorksheet

        set retVal [catch {open $matFileName "r"} matFp]
        if { $retVal != 0 } {
            error "Cannot open file $matFileName"
        }
        fconfigure $matFp -translation binary

        set headerList [::Excel::_GetMatlabHeader $matFp]
        lassign $headerList version width height

        # Parse a Level 4 MAT-File
        if { $version == 4 } {
            set bytesPerPixel 8
            for { set col 0 } { $col < $width } { incr col } {
                for { set row 0 } { $row < $height } { incr row } {
                    set valBytes [read $matFp $bytesPerPixel]
                    binary scan $valBytes d val
                    lappend rowList($row) $val
                }
            }
        }

        for { set row 0 } { $row < $height } { incr row } {
            lappend matrixList $rowList($row)
        }

        close $matFp
        return $matrixList
    }

    proc WriteMatlabFile { matrixList matFileName } {
        # Write the values of a matrix into a Matlab file.
        #
        # matrixList  - Floating point matrix.
        # matFileName - Name of the Matlab file.
        #
        # Note: Only Matlab Level 4 files are currently supported.
        # See SetMatrixValues for the description of a matrix representation.
        #
        # No return value.
        #
        # See also: ReadMatlabFile WorksheetToMatlabFile

        set retVal [catch {open $matFileName "w"} matFp]
        if { $retVal != 0 } {
            error "Cannot open file $matFileName"
        }
        fconfigure $matFp -translation binary

        set height [llength $matrixList]
        set width  [llength [lindex $matrixList 0]]
        ::Excel::_PutMatlabHeader $matFp $width $height [file rootname $matFileName]
        for { set col 0 } { $col < $width } { incr col } {
            for { set row 0 } { $row < $height } { incr row } {
                set pix [lindex [lindex $matrixList $row] $col]
                puts -nonewline $matFp [binary format d $pix]
            }
        }
    }

    proc MatlabFileToWorksheet { matFileName worksheetId { useHeader true } } {
        # Insert the data values of a Matlab file into a worksheet.
        #
        # matFileName - Name of the Matlab file.
        # worksheetId - Identifier of the worksheet.
        # useHeader   - true: Insert the header of the Matlab file as first row.
        #               false: Only transfer the data values as floating point values.
        #
        # The header information are as follows: MatlabVersion Width Height
        #
        # Note: Only Matlab Level 4 files are currently supported.
        #
        # No return value.
        #
        # See also: WorksheetToMatrixValues SetMatrixValues
        # WikitFileToWorksheet MediaWikiFileToWorksheet RawImageFileToWorksheet
        # TablelistToWorksheet WordTableToWorksheet

        set startRow 1
        if { $useHeader } {
            set headerList [::Excel::ReadMatlabHeader $matFileName]
            ::Excel::SetHeaderRow $worksheetId $headerList
            incr startRow
        }
        set matrixList [::Excel::ReadMatlabFile $matFileName]
        ::Excel::SetMatrixValues $worksheetId $matrixList $startRow 1
    }

    proc WorksheetToMatlabFile { worksheetId matFileName { useHeader true } } {
        # Insert the values of a worksheet into a Matlab file.
        #
        # worksheetId - Identifier of the worksheet.
        # matFileName - Name of the Matlab file.
        # useHeader   - true: Interpret the first row of the worksheet as header and
        #                     thus do not transfer this row into the Matlab file.
        #               false: All worksheet cells are interpreted as data.
        #
        # Note: Only Matlab Level 4 files are currently supported.
        #
        # No return value.
        #
        # See also: MatlabFileToWorksheet GetMatrixValues
        # WorksheetToWikitFile WorksheetToMediaWikiFile WorksheetToRawImageFile
        # WorksheetToTablelist WorksheetToWordTable

        set numRows [::Excel::GetLastUsedRow $worksheetId]
        set numCols [::Excel::GetLastUsedColumn $worksheetId]
        set startRow 1
        if { $useHeader } {
            incr startRow
        }
        set excelList [::Excel::GetMatrixValues $worksheetId $startRow 1 $numRows $numCols]
        WriteMatlabFile $excelList $matFileName
    }

    proc MatlabFileToExcelFile { matFileName excelFileName \
                                { useHeader true } { quitExcel true } } {
        # Convert a Matlab table file to an Excel file.
        #
        # matFileName   - Name of the Matlab input file.
        # excelFileName - Name of the Excel output file.
        # useHeader     - true: Insert the header of the Matlab file as first row.
        #                 false: Only transfer the data values as floating point values.
        # quitExcel     - true:  Quit the Excel instance after generation of output file.
        #                 false: Leave the Excel instance open after generation of output file.
        #
        # The table data from the Matlab file will be inserted into a worksheet name "Matlab".
        #
        # Return the Excel application identifier, if quitExcel is false.
        # Otherwise no return value.
        #
        # See also: MatlabFileToWorksheet ExcelFileToMatlabFile ReadMatlabFile

        set appId [::Excel::OpenNew true]
        set workbookId [::Excel::AddWorkbook $appId]
        set worksheetId [::Excel::AddWorksheet $workbookId "Matlab"]
        ::Excel::MatlabFileToWorksheet $matFileName $worksheetId $useHeader
        ::Excel::SaveAs $workbookId $excelFileName
        if { $quitExcel } {
            ::Excel::Quit $appId
        } else {
            return $appId
        }
    }

    proc ExcelFileToMatlabFile { excelFileName matFileName { worksheetNameOrIndex 0 } \
                                { useHeader true } { quitExcel true } } {
        # Convert an Excel file to a Matlab table file.
        #
        # excelFileName        - Name of the Excel input file.
        # matFileName          - Name of the Matlab output file.
        # worksheetNameOrIndex - Worksheet name or index to convert.
        # useHeader            - true:  Use the first row of the worksheet as the header
        #                               of the Matlab file.
        #                        false: Do not generate a Matlab file header. All worksheet
        #                               cells are interpreted as data.
        # quitExcel            - true:  Quit the Excel instance after generation of output file.
        #                        false: Leave the Excel instance open after generation of output file.
        #
        # Note, that the Excel Workbook is opened in read-only mode.
        #
        # Return the Excel application identifier, if quitExcel is false.
        # Otherwise no return value.
        #
        # See also: MatlabFileToWorksheet MatlabFileToExcelFile
        # ReadMatlabFile WriteMatlabFile

        set appId [::Excel::OpenNew true]
        set workbookId [::Excel::OpenWorkbook $appId $excelFileName true]
        if { [string is integer $worksheetNameOrIndex] } {
            set worksheetId [::Excel::GetWorksheetIdByIndex $workbookId [expr int($worksheetNameOrIndex)]]
        } else {
            set worksheetId [::Excel::GetWorksheetIdByName $workbookId $worksheetNameOrIndex]
        }
        ::Excel::WorksheetToMatlabFile $worksheetId $matFileName $useHeader
        if { $quitExcel } {
            ::Excel::Quit $appId
        } else {
            return $appId
        }
    }
}
