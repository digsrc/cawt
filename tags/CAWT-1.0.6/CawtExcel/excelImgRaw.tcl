# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval ::Excel {

    proc _GetRawImageHeader { rawFp } {
        if { [gets $rawFp line] >= 0 } {
            scan $line "Magic=%s" magic
            if { $magic ne "RAW" } {
                error "Invalid Magic value: $magic (must be RAW)"
            }
        } else {
            error "Error while trying to parse Magic keyword"
        }
        if { [gets $rawFp line] >= 0 } {
            scan $line "Width=%d" width
            if { $width <= 0 } {
                error "Invalid Width value: $width (must be greater than zero)"
            }
        } else {
            error "Error while trying to parse Width keyword"
        }
        if { [gets $rawFp line] >= 0 } {
            scan $line "Height=%d" height
            if { $height <= 0 } {
                error "Invalid Height value: $height (must be greater than zero)"
            }
        } else {
            error "Error while trying to parse Height keyword"
        }
        if { [gets $rawFp line] >= 0 } {
            scan $line "NumChan=%d" numChans
            if { $numChans <= 0 || $numChans > 4 } {
                error "Invalid NumChans value: $numChans (must be in 1..4)"
            }
        } else {
            error "Error while trying to parse NumChan keyword"
        }
        if { [gets $rawFp line] >= 0 } {
            scan $line "ByteOrder=%s" byteOrder
            if { $byteOrder ne "Intel" && $byteOrder ne "Motorola" } {
                error "Invalid ByteOrder value: $byteOrder (must be Intel or Motorola)"
            }
        } else {
            error "Error while trying to parse ByteOrder keyword"
        }
        if { [gets $rawFp line] >= 0 } {
            scan $line "ScanOrder=%s" scanOrder
            if { $scanOrder ne "TopDown" && $scanOrder ne "BottomUp" } {
                error "Invalid ScanOrder value: $scanOrder (must be TopDown or BottomUp)"
            }
        } else {
            error "Error while trying to parse ScanOrder keyword"
        }
        if { [gets $rawFp line] >= 0 } {
            scan $line "PixelType=%s" pixelType
            if { $pixelType ne "byte" && $pixelType ne "short" && $pixelType ne "float" } {
                error "Invalid PixelType value: $pixelType (must be byte, short or float)"
            }
        } else {
            error "Error while trying to parse PixelType keyword"
        }
        return [list $magic $width $height $numChans $byteOrder $scanOrder $pixelType]
    }

    proc _PrintHeaderLine { fp msg } {
        set HEADLEN 20
        while { [string length $msg] < [expr {$HEADLEN -1}] } {
            append msg " "
        }
        puts $fp $msg
    }

    proc _GetNativeByteOrder {} {
        if { $::tcl_platform(byteOrder) eq "littleEndian" } {
            return "Intel"
        } else {
            return "Motorola"
        }
    }

    proc _PutRawImageHeader { rawFp width height } {
        ::Excel::_PrintHeaderLine $rawFp [format "Magic=%s" "RAW"]
        ::Excel::_PrintHeaderLine $rawFp [format "Width=%d"  $width]
        ::Excel::_PrintHeaderLine $rawFp [format "Height=%d" $height]
        ::Excel::_PrintHeaderLine $rawFp [format "NumChan=%d" 1]
        ::Excel::_PrintHeaderLine $rawFp [format "ByteOrder=%s" [::Excel::_GetNativeByteOrder]]
        ::Excel::_PrintHeaderLine $rawFp [format "ScanOrder=%s" "TopDown"]
        ::Excel::_PrintHeaderLine $rawFp [format "PixelType=%s" "float"]
    }

    proc ReadRawImageHeader { rawImgFile } {
        # Read the header of a raw photo image.
        #
        # rawImgFile - File name of the image.
        #
        # Return the header information as a list containing the following values:
        # Magic Width Height NumChan ByteOrder ScanOrder PixelType
        #
        # See also: ReadRawImageFile

        set retVal [catch {open $rawImgFile "r"} rawFp]
        if { $retVal != 0 } {
            error "Cannot open file $rawImgFile"
        }
        fconfigure $rawFp -translation binary
        set headerList [::Excel::_GetRawImageHeader $rawFp]
        close $rawFp
        return $headerList
    }

    proc ReadRawImageFile { rawImgFile } {
        # Read a raw photo image into a matrix.
        #
        # rawImgFile - File name of the image.
        #
        # Note: Only 1-channel floating-point raw images are currently supported.
        #
        # Return the image data as a matrix.
        # See SetMatrixValues for the description of a matrix representation.
        #
        # See also: ReadRawImageHeader WriteRawImageFile RawImageFileToWorksheet

        set retVal [catch {open $rawImgFile "r"} rawFp]
        if { $retVal != 0 } {
            error "Cannot open file $rawImgFile"
        }
        fconfigure $rawFp -translation binary

        set headerList [::Excel::_GetRawImageHeader $rawFp]
        lassign $headerList magic width height numChans byteOrder scanOrder pixelType

        if { $numChans != 1 && $pixelType ne "float" } {
            error "Only 1-channel floating point images are currently supported."
        }
        if { $byteOrder eq [::Excel::_GetNativeByteOrder] } {
            set scanFmt "f"
        } elseif { $byteOrder eq "Intel" } {
            set scanFmt "r"
        } elseif { $byteOrder eq "Motorola" } {
            set scanFmt "R"
        } else {
            error "Invalid byte order \"$byteOrder\" in file."
        }

        set numVals [expr {$width*$height}]
        for { set row 0 } { $row < $height } { incr row } {
            for { set col 0 } { $col < $width } { incr col } {
                set valBytes [read $rawFp 4]
                binary scan $valBytes $scanFmt val
                lappend rowList($row) $val
            }
        }

        if { $scanOrder eq "TopDown" } {
            for { set row 0 } { $row < $height } { incr row } {
                lappend matrixList $rowList($row)
            }
        } else {
            for { set row [expr $height-1] } { $row >= 0 } { incr row -1 } {
                lappend matrixList $rowList($row)
            }
        }

        close $rawFp
        return $matrixList
    }

    proc WriteRawImageFile { matrixList rawImgFile } {
        # Write the values of a matrix into a raw photo image file.
        #
        # matrixList - Floating point matrix.
        # rawImgFile - File name of the image.
        #
        # Note: The matrix values are written as a 1-channel floating-point image.
        #
        # See SetMatrixValues for the description of a matrix representation.
        #
        # No return value.
        #
        # See also: ReadRawImageFile WorksheetToRawImageFile

        set retVal [catch {open $rawImgFile "w"} rawFp]
        if { $retVal != 0 } {
            error "Cannot open file $rawImgFile"
        }
        fconfigure $rawFp -translation binary

        set height [llength $matrixList]
        set width  [llength [lindex $matrixList 0]]
        ::Excel::_PutRawImageHeader $rawFp $width $height
        foreach rowList $matrixList {
            foreach pix $rowList {
                puts -nonewline $rawFp [binary format f $pix]
            }
        }
    }

    proc RawImageFileToWorksheet { rawFileName worksheetId { useHeader true } } {
        # Insert the pixel values of a raw photo image into a worksheet.
        #
        # rawFileName - File name of the image.
        # worksheetId - Identifier of the worksheet.
        # useHeader   - true: Insert the header of the raw image as first row.
        #               false: Only transfer the pixel values as floating point values.
        #
        # The header information are as follows:
        # Magic Width Height NumChan ByteOrder ScanOrder PixelType
        #
        # Note: Only 1-channel floating-point raw images are currently supported.
        #
        # No return value.
        #
        # See also: WorksheetToRawImageFile SetMatrixValues
        # WikitFileToWorksheet MediaWikiFileToWorksheet MatlabFileToWorksheet
        # TablelistToWorksheet WordTableToWorksheet

        set startRow 1
        if { $useHeader } {
            set headerList [::Excel::ReadRawImageHeader $rawFileName]
            ::Excel::SetHeaderRow $worksheetId $headerList
            incr startRow
        }
        set matrixList [::Excel::ReadRawImageFile $rawFileName]
        ::Excel::SetMatrixValues $worksheetId $matrixList $startRow 1
    }

    proc WorksheetToRawImageFile { worksheetId rawFileName { useHeader true } } {
        # Insert the values of a worksheet into a raw photo image file.
        #
        # worksheetId - Identifier of the worksheet.
        # rawFileName - File name of the image.
        # useHeader   - true: Interpret the first row of the worksheet as header and
        #                     thus do not transfer this row into the image.
        #               false: All worksheet cells are interpreted as data.
        #
        # The image generated is a 1-channel floating point photo image. It can be
        # read and manipulated with the Img extension. It is not a "raw" image as used
        # with digital cameras, but just "raw" image data.
        #
        # No return value.
        #
        # See also: RawImageFileToWorksheet GetMatrixValues
        # WorksheetToWikitFile WorksheetToMediaWikiFile WorksheetToMatlabFile
        # WorksheetToTablelist WorksheetToWordTable

        set numRows [::Excel::GetLastUsedRow $worksheetId]
        set numCols [::Excel::GetLastUsedColumn $worksheetId]
        set startRow 1
        if { $useHeader } {
            incr startRow
        }
        set excelList [::Excel::GetMatrixValues $worksheetId $startRow 1 $numRows $numCols]
        WriteRawImageFile $excelList $rawFileName
    }

    proc RawImageFileToExcelFile { rawFileName excelFileName \
                                   { useHeader true } { quitExcel true } } {
        # Convert a raw photo image file to an Excel file.
        #
        # rawFileName   - Name of the raw photo image input file.
        # excelFileName - Name of the Excel output file.
        # useHeader     - true:  Use header information from the image file to
        #                        generate an Excel header (see SetHeaderRow).
        #                 false: Only transfer the image data.
        # quitExcel     - true:  Quit the Excel instance after generation of output file.
        #                 false: Leave the Excel instance open after generation of output file.
        #
        # The table data from the image file will be inserted into a worksheet name "RawImage".
        #
        # Return the Excel application identifier, if quitExcel is false.
        # Otherwise no return value.
        #
        # See also: RawImageFileToWorksheet ExcelFileToRawImageFile ReadRawImageFile WriteRawImageFile

        set appId [::Excel::OpenNew true]
        set workbookId [::Excel::AddWorkbook $appId]
        set worksheetId [::Excel::AddWorksheet $workbookId "RawImage"]
        ::Excel::RawImageFileToWorksheet $rawFileName $worksheetId $useHeader
        ::Excel::SaveAs $workbookId $excelFileName
        if { $quitExcel } {
            ::Excel::Quit $appId
        } else {
            return $appId
        }
    }

    proc ExcelFileToRawImageFile { excelFileName rawFileName { worksheetNameOrIndex 0 } \
                                  { useHeader true } { quitExcel true } } {
        # Convert an Excel file to a raw photo image file.
        #
        # excelFileName        - Name of the Excel input file.
        # rawFileName          - Name of the image output file.
        # worksheetNameOrIndex - Worksheet name or index to convert.
        # useHeader            - true:  Use the first row of the worksheet as the header
        #                               of the raw image file.
        #                        false: Do not generate a raw image file header. All worksheet
        #                               cells are interpreted as data.
        # quitExcel             - true:  Quit the Excel instance after generation of output file.
        #                         false: Leave the Excel instance open after generation of output file.
        #
        # Note, that the Excel Workbook is opened in read-only mode.
        #
        # Return the Excel application identifier, if quitExcel is false.
        # Otherwise no return value.
        #
        # See also: RawImageFileToWorksheet RawImageFileToExcelFile ReadRawImageFile WriteRawImageFile

        set appId [::Excel::OpenNew true]
        set workbookId [::Excel::OpenWorkbook $appId $excelFileName true]
        if { [string is integer $worksheetNameOrIndex] } {
            set worksheetId [::Excel::GetWorksheetIdByIndex $workbookId [expr int($worksheetNameOrIndex)]]
        } else {
            set worksheetId [::Excel::GetWorksheetIdByName $workbookId $worksheetNameOrIndex]
        }
        ::Excel::WorksheetToRawImageFile $worksheetId $rawFileName $useHeader
        if { $quitExcel } {
            ::Excel::Quit $appId
        } else {
            return $appId
        }
    }
}
