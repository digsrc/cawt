# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval ::Excel {

    variable excelVersion     "0.0"
    variable excelAppName     "Excel.Application"
    variable decimalSeparator ""
    variable _ruffdoc

    lappend _ruffdoc Introduction {
        The Excel namespace provides commands to control Microsoft Excel.
    }

    proc _setFloatSeparator { appId } {
        variable excelVersion
        variable decimalSeparator

        # Method DecimalSeparator is only available in Excel 2003 and up.
        if { $excelVersion >= 11.0 } {
            set decimalSeparator [$appId DecimalSeparator]
        } else {
            # Set variable decimalSeparator to ",", if using German versions
            # of Excel older than 2003.
            # Note, that these older versions are not supported anymore.
            set decimalSeparator "."
            # set decimalSeparator ","
        }
    }

    proc GetFloatSeparator {} {
        # Return the decimal separator used by Excel.
        #
        # Only valid, after a call of Open or OpenNew.
        #
        # See also: GetVersion

        variable decimalSeparator

        return $decimalSeparator
    }

    proc GetLangNumberFormat { pre post } {
        # Return an Excel number format string.
        #
        # pre  - Number of digits before the decimal point.
        # post - Number of digits after the decimal point.
        #
        # The number of digits is specified as a string containing as
        # many zeros as wanted digits.
        #
        # Example: [GetLangNumberFormat "0" "0000"] will return the Excel format string to show
        #          floating point values with 4 digits after the decimal point.
        #
        # See also: SetRangeFormat

        set floatSep [::Excel::GetFloatSeparator]
        return [format "%s%s%s" $pre $floatSep $post]
    }

    proc ColumnCharToInt { colChar } {
        # Return an Excel column string as a column number.
        #
        # colChar - Column string.
        #
        # Example: [::Excel::ColumnCharToInt A] returns 1.
        #          [::Excel::ColumnCharToInt Z] returns 26.
        #
        # See also: ColumnIntToChar

        set abc {- A B C D E F G H I J K L M N O P Q R S T U V W X Y Z}
        set int 0
        foreach char [split $colChar ""] {
            set int [expr {$int*26 + [lsearch $abc $char]}]
        }
        return $int
    }

    proc ColumnIntToChar { col } {
        # Return a column number as an Excel column string.
        #
        # col - Column number.
        #
        # Example: [::Excel::ColumnIntToChar 1]  returns "A".
        #          [::Excel::ColumnIntToChar 26] returns "Z".
        #
        # See also: ColumnCharToInt

        if { $col <= 0 } {
            error "Column number $col is invalid."
        }
        set dividend $col
        set columnName ""
        while { $dividend > 0 } {
            set modulo [expr { ($dividend - 1) % 26 } ]
            set columnName [format "%c${columnName}" [expr { 65 + $modulo} ] ]
            set dividend [expr { ($dividend - $modulo) / 26 } ]
        }
        return $columnName
    }

    proc GetCellRange { row1 col1 row2 col2 } {
        # Return a numeric cell range as an Excel range string.
        #
        # row1 - Row number of upper-left corner of the cell range.
        # col1 - Column number of upper-left corner of the cell range.
        # row2 - Row number of lower-right corner of the cell range.
        # col2 - Column number of lower-right corner of the cell range.
        #
        # Example: [GetCellRange 1 2  5 7] returns string "B1:G5".
        #
        # See also: GetColumnRange

        set range [format "%s%d:%s%d" \
                   [ColumnIntToChar $col1] $row1 \
                   [ColumnIntToChar $col2] $row2]
        return $range
    }

    proc GetColumnRange { col1 col2 } {
        # Return a numeric column range as an Excel range string.
        #
        # col1 - Column number of the left-most column.
        # col2 - Column number of the right-most column.
        #
        # Example: [GetColumnRange 2 7] returns string "B:G".
        #
        # See also: GetCellRange

        set range [format "%s:%s" \
                   [ColumnIntToChar $col1] \
                   [ColumnIntToChar $col2]]
        return $range
    }

    proc GetNumRows { rangeId } {
        # Return the number of rows of a cell range.
        #
        # rangeId - Identifier of a range, cells collection or a worksheet.
        #
        # If specifying a worksheetId or cellsId, the maximum number of rows
        # of a worksheet will be returned.
        # The maximum number of rows is 65.536 for Excel versions before 2007.
        # Since 2007 the maximum number of rows is 1.048.576.
        #
        # See also: GetNumUsedRows GetFirstUsedRow GetLastUsedRow GetNumColumns

        return [$rangeId -with { Rows } Count]
    }

    proc GetNumUsedRows { worksheetId } {
        # Return the number of used rows of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        #
        # See also: GetNumRows GetFirstUsedRow GetLastUsedRow GetNumUsedColumns

        return [$worksheetId -with { UsedRange Rows } Count]
    }

    proc GetFirstUsedRow { worksheetId } {
        # Return the index of the first used row of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        #
        # See also: GetNumRows GetNumUsedRows GetLastUsedRow GetNumUsedColumns

        return [$worksheetId -with { UsedRange } Row]
    }

    proc GetLastUsedRow { worksheetId } {
        # Return the index of the last used row of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        #
        # See also: GetNumRows GetNumUsedRows GetFirstUsedRow GetNumUsedColumns

        return [expr { [::Excel::GetFirstUsedRow $worksheetId] + \
                       [::Excel::GetNumUsedRows $worksheetId] - 1 }]
    }

    proc GetNumColumns { rangeId } {
        # Return the number of columns of a cell range.
        #
        # rangeId - Identifier of a range, cells collection or a worksheet.
        #
        # If specifying a worksheetId or cellsId, the maximum number of columns
        # of a worksheet will be returned.
        # The maximum number of columns is 256 for Excel versions before 2007.
        # Since 2007 the maximum number of columns is 16.384.
        #
        # See also: GetNumUsedColumns GetFirstUsedColumn GetLastUsedColumn GetNumRows

        return [$rangeId -with { Columns } Count]
    }

    proc GetNumUsedColumns { worksheetId } {
        # Return the number of used columns of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        #
        # In some cases the number of columns returned may be 1 to high.
        #
        # See also: GetNumColumns GetFirstUsedColumn GetLastUsedColumn GetNumUsedRows

        return [$worksheetId -with { UsedRange Columns } Count]
    }

    proc GetFirstUsedColumn { worksheetId } {
        # Return the index of the first used column of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        #
        # See also: GetNumColumns GetNumUsedColumns GetLastUsedColumn GetNumUsedRows
 
        return [$worksheetId -with { UsedRange } Column]
    }

    proc GetLastUsedColumn { worksheetId } {
        # Return the index of the last used column of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        #
        # See also: GetNumColumns GetNumUsedColumns GetFirstUsedColumn GetNumUsedRows

        return [expr { [::Excel::GetFirstUsedColumn $worksheetId] + \
                       [::Excel::GetNumUsedColumns $worksheetId] - 1 }]
    }

    proc SelectRangeByString { worksheetId rangeStr { visSel false } } {
        # Select a range by specifying an Excel range string.
        #
        # worksheetId - Identifier of the worksheet.
        # rangeStr    - String specifying a cell range.
        # visSel      - true: See the selection in the user interface.
        #               false: The selection is not visible.
        #
        # Return the range identifier of the cell range.
        #
        # See also: SelectRangeByIndex GetCellRange

        set cellsId [::Excel::GetCellsId $worksheetId]
        set rangeId [$cellsId Range $rangeStr]
        if { $visSel } {
            $rangeId Select
        }
        ::Cawt::Destroy $cellsId
        return $rangeId
    }

    proc SelectRangeByIndex { worksheetId row1 col1 row2 col2 { visSel false } } {
        # Select a range by specifying a numeric cell range.
        #
        # worksheetId - Identifier of the worksheet.
        # row1        - Row number of upper-left corner of the cell range.
        # col1        - Column number of upper-left corner of the cell range.
        # row2        - Row number of lower-right corner of the cell range.
        # col2        - Column number of lower-right corner of the cell range.
        # visSel      - true: See the selection in the user interface.
        #               false: The selection is not visible.
        #
        # Return the range identifier of the cell range.
        #
        # See also: SelectCellByIndex GetCellRange

        set rangeStr [::Excel::GetCellRange $row1 $col1 $row2 $col2]
        return [::Excel::SelectRangeByString $worksheetId $rangeStr $visSel]
    }

    proc SelectAll { worksheetId } {
        # Select all cells of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        #
        # Return the cells collection of the selected cells.
        #
        # See also: CopyWorksheet

        return [[::Cawt::GetApplicationId $worksheetId] Cells]
    }

    proc GetRangeCharacters { rangeId { start 1 } { length -1 } } {
        # Return characters of a cell range.
        #
        # rangeId - Identifier of the cell range.
        # start   - Start of the character range.
        # length  - The number of characters after start.
        #
        # Return all or a range of characters of a cell range.
        # If no optional parameters are specified, all characters of the cell range are
        # returned.
        #
        # See also: SelectRangeByIndex SelectRangeByString

        if { $length < 0 } {
            return [$rangeId Characters $start]
        } else {
            return [$rangeId Characters $start $length]
        }
    }

    proc SetRangeFontSubscript { rangeId { onOff true } } {
        # Set the subscript font style of a cell or character range.
        #
        # rangeId - Identifier of the cell range.
        # onOff   - true: Set subscript style on.
        #           false: Set subscript style off.
        #
        # No return value.
        #
        # See also: SetRangeFontSuperscript SelectRangeByIndex GetRangeCharacters

        $rangeId -with { Font } Subscript [::Cawt::TclBool $onOff]
    }

    proc SetRangeFontSuperscript { rangeId { onOff true } } {
        # Set the superscript font style of a cell or character range.
        #
        # rangeId - Identifier of the cell range.
        # onOff   - true: Set superscript style on.
        #           false: Set superscript style off.
        #
        # No return value.
        #
        # See also: SetRangeFontSubscript SelectRangeByIndex GetRangeCharacters

        $rangeId -with { Font } Superscript [::Cawt::TclBool $onOff]
    }

    proc SetRangeFontBold { rangeId { onOff true } } {
        # Set the bold font style of a cell range.
        #
        # rangeId - Identifier of the cell range.
        # onOff   - true: Set bold style on.
        #           false: Set bold style off.
        #
        # No return value.
        #
        # See also: SetRangeFontItalic SelectRangeByIndex SelectRangeByString

        $rangeId -with { Font } Bold [::Cawt::TclBool $onOff]
    }

    proc SetRangeFontItalic { rangeId { onOff true } } {
        # Set the italic font style of a cell range.
        #
        # rangeId - Identifier of the cell range.
        # onOff   - true: Set italic style on.
        #           false: Set italic style off.
        #
        # No return value.
        #
        # See also: SetRangeFontBold SelectRangeByIndex SelectRangeByString

        $rangeId -with { Font } Italic [::Cawt::TclBool $onOff]
    }

    proc SetRangeMergeCells { rangeId { onOff true } } {
        # Merge/Unmerge a range of cells.
        #
        # rangeId - Identifier of the cell range.
        # onOff   - true: Set cell merge on.
        #           false: Set cell merge off.
        #
        # No return value.
        #
        # See also: SetRangeVerticalAlignment SelectRangeByIndex SelectRangeByString

        $rangeId MergeCells [::Cawt::TclBool $onOff]
    }

    proc SetRangeHorizontalAlignment { rangeId align } {
        # Set the horizontal alignment of a cell range.
        #
        # rangeId - Identifier of the cell range.
        # align   - Value of enumeration type XlHAlign (see excelConst.tcl).
        #
        # No return value.
        #
        # See also: SetRangeVerticalAlignment SelectRangeByIndex SelectRangeByString

        $rangeId HorizontalAlignment [expr $align]
    }

    proc SetRangeVerticalAlignment { rangeId align } {
        # Set the vertical alignment of a cell range.
        #
        # rangeId - Identifier of the cell range.
        # align   - Value of enumeration type XlVAlign (see excelConst.tcl).
        #
        # No return value.
        #
        # See also: SetRangeHorizontalAlignment SelectRangeByIndex SelectRangeByString

        $rangeId VerticalAlignment [expr $align]
    }

    proc ToggleAutoFilter { rangeId } {
        # Toggle the AutoFilter switch of a cell range.
        #
        # rangeId - Identifier of the cell range.
        #
        # No return value.
        #
        # See also: SelectRangeByIndex SelectRangeByString

        $rangeId AutoFilter
    }

    proc SetRangeFillColor { rangeId r g b } {
        # Set the fill color of a cell range.
        #
        # rangeId - Identifier of the cell range.
        # r       - Red component of the text color.
        # g       - Green component of the text color.
        # b       - Blue component of the text color.
        #
        # The r, g and b values are specified as integers in the
        # range [0, 255].
        #
        # No return value.
        #
        # See also: SetRangeTextColor ::Cawt::RgbToColor SelectRangeByIndex SelectRangeByString

        set color [::Cawt::RgbToColor $r $g $b]
        $rangeId -with { Interior } Color $color
        $rangeId -with { Interior } Pattern $::Excel::xlSolid
    }

    proc SetRangeTextColor { rangeId r g b } {
        # Set the text color of a cell range.
        #
        # rangeId - Identifier of the cell range.
        # r       - Red component of the text color.
        # g       - Green component of the text color.
        # b       - Blue component of the text color.
        #
        # The r, g and b values are specified as integers in the
        # range [0, 255].
        #
        # No return value.
        #
        # See also: SetRangeFillColor ::Cawt::RgbToColor SelectRangeByIndex SelectRangeByString

        set color [::Cawt::RgbToColor $r $g $b]
        $rangeId -with { Font } Color $color
    }

    proc SetRangeBorder { rangeId side \
                          { weight $::Excel::xlThin } \
                          { lineStyle $::Excel::xlContinuous } \
                          { r 0 } { g 0 } { b 0 } } {
        # Set the attributes of one border of a cell range.
        #
        # rangeId   - Identifier of the cell range.
        # side      - Value of enumeration type XlBordersIndex (see excelConst.tcl).
        #             Typical values: xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight.
        # weight    - Value of enumeration type XlBorderWeight (see excelConst.tcl).
        #             Typical values: xlThin, xlMedium, xlThick.
        # lineStyle - Value of enumeration type XlLineStyle (see excelConst.tcl).
        #             Typical values: xlContinuous, xlDash, xlDot.
        # r         - Red component of the border color.
        # g         - Green component of the border color.
        # b         - Blue component of the border color.
        #
        # The r, g and b values are specified as integers in the
        # range [0, 255].
        #
        # No return value.
        #
        # See also: SetRangeBorders SelectRangeByIndex SelectRangeByString

        set color [::Cawt::RgbToColor $r $g $b]
        set borders [$rangeId Borders]
        [$borders Item $side] Weight    [expr $weight]
        [$borders Item $side] LineStyle [expr $lineStyle]
        [$borders Item $side] Color     [expr $color]
    }

    proc SetRangeBorders { rangeId \
                          { weight $::Excel::xlThin } \
                          { lineStyle $::Excel::xlContinuous } \
                          { r 0 } { g 0 } { b 0 } } {
        # Set the attributes of all borders of a cell range.
        #
        # rangeId   - Identifier of the cell range.
        # weight    - Value of enumeration type XlBorderWeight (see excelConst.tcl).
        #             Typical values: xlThin, xlMedium, xlThick.
        # lineStyle - Value of enumeration type XlLineStyle (see excelConst.tcl).
        #             Typical values: xlContinuous, xlDash, xlDot.
        # r         - Red component of the border color.
        # g         - Green component of the border color.
        # b         - Blue component of the border color.
        #
        # The r, g and b values are specified as integers in the
        # range [0, 255].
        #
        # No return value.
        #
        # See also: SetRangeBorder SelectRangeByIndex SelectRangeByString

        ::Excel::SetRangeBorder $rangeId $::Excel::xlEdgeLeft   $weight $lineStyle $r $g $b
        ::Excel::SetRangeBorder $rangeId $::Excel::xlEdgeRight  $weight $lineStyle $r $g $b
        ::Excel::SetRangeBorder $rangeId $::Excel::xlEdgeBottom $weight $lineStyle $r $g $b
        ::Excel::SetRangeBorder $rangeId $::Excel::xlEdgeTop    $weight $lineStyle $r $g $b
    }

    proc SetRangeFormat { rangeId fmt { subFmt "" } } {
        # Set the format of a cell range.
        #
        # rangeId - Identifier of the cell range.
        # fmt     - Format of the cell range.  Possible values: "text", "int", "real".
        # subFmt  - Sub-format of the cell range. Only valid, if fmt is "real". Then it
        #           specifies the number of digits before and after the decimal point.
        #           Use the GetLangNumberFormat procedure for specifying the sub-format.
        #           If subFmt is the empty string 2 digits after the decimal point are used.
        #
        # No return value.
        #
        # See also: SelectRangeByIndex SelectRangeByString

        if { $fmt eq "text" } {
            $rangeId NumberFormat "@"
        } elseif { $fmt eq "int" } {
            $rangeId NumberFormat "0"
        } elseif { $fmt eq "real" } {
            if { $subFmt eq "" } {
                set subFmt [::Excel::GetLangNumberFormat "0" "00"]
            }
            $rangeId NumberFormat $subFmt
        } else {
            error "Invalid cell format \"$fmt\" given"
        }
    }

    proc SetCommentDisplayMode { appId { showComment false } { showIndicator true } } {
        # Set the global display mode of comments.
        #
        # appId         - Identifier of the Excel instance.
        # showComment   - true:  Show the comments.
        #                 false: Do not show the comments.
        # showIndicator - true:  Show an indicator for the comments.
        #                 false: Do not show an indicator.
        #
        # No return value.
        #
        # See also: SetRangeComment

        if { $showComment && $showIndicator } {
            $appId DisplayCommentIndicator $::Excel::xlCommentAndIndicator
        } elseif { $showIndicator } {
            $appId DisplayCommentIndicator $::Excel::xlCommentIndicatorOnly
        } else {
            $appId DisplayCommentIndicator $::Excel::xlNoIndicator
        }
    }

    proc SetRangeComment { rangeId comment { imgFileName "" } { addUserName true } { visible false } } {
        # Set the comment text of a cell range.
        #
        # rangeId     - Identifier of the cell range.
        # comment     - Comment text.
        # imgFileName - File name of an image used as comment background (as absolute path).
        # addUserName - Automatically add user name before comment text.
        # visible     - true: Show the comment window.
        #               false: Hide the comment window.
        #
        # Note, that an already existing comment is overwritten.
        #
        # Return the comment identifier.
        #
        # See also: SelectRangeByIndex SelectRangeByString SetCommentDisplayMode ::Cawt::GetUserName

        set commentId [$rangeId Comment]
        if { ! [::Cawt::IsValidId $commentId] } {
            set commentId [$rangeId AddComment]
        }
        $commentId Visible [::Cawt::TclBool $visible]
        if { $addUserName } {
            set userName [::Cawt::GetUserName [$commentId Application]]
            set msg [format "%s:\n%s" $userName $comment]
        } else {
            set msg $comment
        }
        $commentId Text $msg
        if { $imgFileName ne "" } {
            set fileName [file nativename $imgFileName]
            $commentId -with { Shape Fill } UserPicture $fileName
        }
        return $commentId
    }

    proc SetCommentSize { commentId width height } {
        # Set the shape size of a comment.
        #
        # commentId - Identifier of the comment.
        # width     - Width of the comment.
        # height    - Height of the comment.
        #
        # The size values must be specified in points.
        # Use ::Cawt::CentiMetersToPoints or ::Cawt::InchesToPoints
        # for conversion.
        #
        # No return value.
        #
        # See also: SetRangeComment ::Cawt::CentiMetersToPoints ::Cawt::InchesToPoints
        
        $commentId -with { Shape } LockAspectRatio [::Cawt::TclInt 0]
        $commentId -with { Shape } Height [expr double ($width)]
        $commentId -with { Shape } Width  [expr double ($height)]
    }
    
    proc GetVersion { appId { useString false } } {
        # Return the version of an Excel application.
        #
        # appId     - Identifier of the Excel instance.
        # useString - true: Return the version name (ex. "Excel 2000").
        #             false: Return the version number (ex. "9.0").
        #
        # Both version name and version number are returned as strings.
        # Version number is in a format, so that it can be evaluated as a
        # floating point number.
        #
        # See also: GetFloatSeparator

        array set map {
            "8.0"  "Excel 97"
            "9.0"  "Excel 2000"
            "10.0" "Excel 2002"
            "11.0" "Excel 2003"
            "12.0" "Excel 2007"
            "14.0" "Excel 2010"
            "15.0" "Excel 2013"
        }
        set version [$appId Version]
        if { $useString } {
            if { [info exists map($version)] } {
                return $map($version)
            } else {
                return "Unknown Excel version"
            }
        } else {
            return $version
        }
        return $version
    }

    proc GetExtString { appId } {
        # Return the default extension of an Excel file.
        #
        # appId - Identifier of the Excel instance.
        #
        # Starting with Excel 12 (2007) this is the string ".xlsx".
        # In previous versions it was ".xls".

        variable excelVersion

        if { $excelVersion >= 12.0 } {
            return ".xlsx"
        } else {
            return ".xls"
        }
    }

    proc OpenNew { { visible true } { width -1 } { height -1 } } {
        # Open a new Excel instance.
        #
        # visible - true: Show the application window.
        #           false: Hide the application window.
        # width   - Width of the application window. If negative, open with last used width.
        # height  - Height of the application window. If negative, open with last used height.
        #
        # Return the identifier of the new Excel application instance.
        #
        # See also: Open Quit Visible

        variable excelAppName
        variable excelVersion

        set appId [::Cawt::GetOrCreateApp $excelAppName false]
        set excelVersion [::Excel::GetVersion $appId]
        ::Excel::_setFloatSeparator $appId
        ::Excel::Visible $appId $visible
        if { $width >= 0 } {
            $appId Width [expr $width]
        }
        if { $height >= 0 } {
            $appId Height [expr $height]
        }
        return $appId
    }

    proc Open { { visible true } { width -1 } { height -1 } } {
        # Open an Excel instance. Use an already running Excel, if available.
        #
        # visible - true: Show the application window.
        #           false: Hide the application window.
        # width   - Width of the application window. If negative, open with last used width.
        # height  - Height of the application window. If negative, open with last used height.
        #
        # Return the identifier of the Excel application instance.
        #
        # See also: OpenNew Quit Visible

        variable excelAppName
        variable excelVersion

        set appId [::Cawt::GetOrCreateApp $excelAppName true]
        set excelVersion [::Excel::GetVersion $appId]
        ::Excel::_setFloatSeparator $appId
        ::Excel::Visible $appId $visible
        if { $width >= 0 } {
            $appId Width [expr $width]
        }
        if { $height >= 0 } {
            $appId Height [expr $height]
        }
        return $appId
    }

    proc Quit { appId { showAlert true } } {
        # Quit an Excel application instance.
        #
        # appId     - Identifier of the Excel instance.
        # showAlert - true: Show an alert window, if there are unsaved changes.
        #             false: Quit without asking and saving any changes.
        #
        # No return value.
        #
        # See also: Open OpenNew

        if { ! $showAlert } {
            ::Cawt::ShowAlerts $appId false
        }
        $appId Quit
    }

    proc Visible { appId visible } {
        # Toggle the visibilty of an Excel application window.
        #
        # appId   - Identifier of the Excel instance.
        # visible - true: Show the application window.
        #           false: Hide the application window.
        #
        # No return value.
        #
        # See also: Open OpenNew SetWindowState ArrangeWindows

        $appId Visible [::Cawt::TclInt $visible]
    }

    proc SetWindowState { appId { windowState $::Excel::xlNormal } } {
        # Set the window state of an Excel application.
        #
        # appId       - Identifier of the Excel instance.
        # windowState - Value of enumeration type XlWindowState (see excelConst.tcl).
        #               Typical values are: xlMaximized, xlMinimized, xlNormal.
        #
        # No return value.
        #
        # See also: Open Visible ArrangeWindows

        $appId -with { Application } WindowState [expr $windowState]
    }

    proc ArrangeWindows { appId { arrangeStyle $::Excel::xlArrangeStyleVertical } } {
        # Arrange the windows of an Excel application.
        #
        # appId        - Identifier of the Excel instance.
        # arrangeStyle - Value of enumeration type XlArrangeStyle (see excelConst.tcl).
        #                Typical values are: xlArrangeStyleHorizontal,
        #                xlArrangeStyleTiled, xlArrangeStyleVertical
        #
        # No return value.
        #
        # See also: Open Visible SetWindowState

        $appId -with { Windows } Arrange [expr $arrangeStyle]
    }

    proc Close { workbookId } {
        # Close a workbook without saving changes.
        #
        # workbookId - Identifier of the workbook.
        #
        # Use the SaveAs method before closing, if you want to save changes.
        #
        # No return value.
        #
        # See also: SaveAs OpenWorkbook

        $workbookId Close [::Cawt::TclBool false]
    }

    proc SaveAs { workbookId fileName { fmt "" } { backup false } } {
        # Save a workbook to an Excel file.
        #
        # workbookId - Identifier of the workbook to save.
        # fileName   - Name of the Excel file.
        # fmt        - Value of enumeration type XlSheetType (see excelConst.tcl).
        #              If not given or the empty string, the file is stored in the native
        #              format corresponding to the used Excel version.
        # backup     - true: Create a backup file before saving.
        #              false: Do not create a backup file.
        #
        # No return value.
        #
        # See also: SaveAsCsv Close OpenWorkbook

        set fileName [file nativename $fileName]
        set appId [::Cawt::GetApplicationId $workbookId]
        ::Cawt::ShowAlerts $appId false
        if { $fmt eq "" } {
            $workbookId SaveAs $fileName
        } else {
            # SaveAs([Filename], [FileFormat], [Password],
            # [WriteResPassword], [ReadOnlyRecommended], [CreateBackup],
            # [AccessMode As XlSaveAsAccessMode = xlNoChange],
            # [ConflictResolution], [AddToMru], [TextCodepage],
            # [TextVisualLayout], [Local])
            $workbookId -callnamedargs SaveAs \
                        FileName $fileName \
                        FileFormat [expr $fmt] \
                        CreateBackup [::Cawt::TclInt $backup]
        }
        ::Cawt::ShowAlerts $appId true
    }

    proc SaveAsCsv { workbookId worksheetId fileName } {
        # Save a worksheet to file in CSV format.
        #
        # workbookId  - Identifier of the workbook containing the worksheet.
        # worksheetId - Identifier of the worksheet to save.
        # fileName    - Name of the CSV file.
        #
        # No return value.
        #
        # See also: SaveAs Close OpenWorkbook

        set fileName [file nativename $fileName]
        set appId [::Cawt::GetApplicationId $workbookId]
        ::Cawt::ShowAlerts $appId false
        # SaveAs(Filename As String, [FileFormat], [Password],
        # [WriteResPassword], [ReadOnlyRecommended], [CreateBackup],
        # [AddToMru], [TextCodepage], [TextVisualLayout], [Local])
        $worksheetId -callnamedargs SaveAs \
                     Filename $fileName \
                     FileFormat $::Excel::xlCSV
        ::Cawt::ShowAlerts $appId true
    }

    proc AddWorkbook { appId { type $::Excel::xlWorksheet } } {
        # Add a new workbook with 1 worksheet.
        #
        # appId - Identifier of the Excel instance.
        # type  - Value of enumeration type XlSheetType (see excelConst.tcl).
        #         Possible values: xlChart, xlDialogSheet, xlExcel4IntlMacroSheet,
        #         xlExcel4MacroSheet, xlWorksheet
        #
        # Return the identifier of the new workbook.
        #
        # See also: OpenWorkbook Close SaveAs

        return [$appId -with { Workbooks } Add [expr $type]]
    }

    proc OpenWorkbook { appId fileName { readOnly false } } {
        # Open a workbook, i.e load an Excel file.
        #
        # appId    - Identifier of the Excel instance.
        # fileName - Name of the Excel file (as absolute path).
        # readOnly - true: Open the workbook in read-only mode.
        #            false: Open the workbook in read-write mode.
        #
        # Return the identifier of the opened workbook. If the workbook was already open,
        # activate that workbook and return the identifier to that workbook.
        #
        # See also: AddWorkbook Close SaveAs

        set nativeName  [file nativename $fileName]
        set workbooks [$appId Workbooks]
        set retVal [catch {[$workbooks Item [file tail $fileName]] Activate} d]
        if { $retVal == 0 } {
            puts "$nativeName already open"
            set workbookId [$workbooks Item [file tail $fileName]]
        } else {
            # Open(Filename As String, [UpdateLinks], [ReadOnly], [Format],
            # [Password], [WriteResPassword], [IgnoreReadOnlyRecommended],
            # [Origin], [Delimiter], [Editable], [Notify], [Converter],
            # [AddToMru], [Local], [CorruptLoad])
            set workbookId [$workbooks -callnamedargs Open \
                                       Filename $nativeName \
                                       ReadOnly [::Cawt::TclInt $readOnly]]
        }
        ::Cawt::Destroy $workbooks
        return $workbookId
    }

    proc GetWorkbookName { workbookId } {
        # Return the name of a workbook.
        #
        # workbookId - Identifier of the workbook.
        #
        # See also: AddWorkbook

        return [$workbookId Name]
    }

    proc GetActiveWorkbook { appId } {
        # Return the active workbook of an application.
        #
        # appId - Identifier of the Excel instance.
        #
        # Return the identifier of the active workbook.
        #
        # See also: OpenWorkbook

        return [$appId ActiveWorkbook]
    }

    proc IsWorkbookProtected { workbookId } {
        # Check, if a workbook is protected.
        #
        # workbookId - Identifier of the workbook to be checked.
        #
        # Return true, if the workbook is protected, otherwise false.
        #
        # See also: OpenWorkbook

        if { [$workbookId ProtectWindows] } {
            return true
        } else {
            return false
        }
    }

    proc AddWorksheet { workbookId name { visibleType $::Excel::xlSheetVisible } } {
        # Add a new worksheet to the end of a workbook.
        #
        # workbookId  - Identifier of the workbook containing the new worksheet.
        # name        - Name of the new worksheet.
        # visibleType - Value of enumeration type XlSheetVisibility (see excelConst.tcl).
        #               Possible values: xlSheetVisible, xlSheetHidden, xlSheetVeryHidden
        #
        # Return the identifier of the new worksheet.
        #
        # See also: GetNumWorksheets DeleteWorksheet

        set worksheets [$workbookId Worksheets]
        set lastWorksheet [$worksheets Item [$worksheets Count]]
        set worksheetId [$worksheets Add]
        $worksheetId Name $name
        $worksheetId Visible [expr $visibleType]
        ::Cawt::Destroy $worksheets
        return $worksheetId
    }

    proc DeleteWorksheet { workbookId worksheetId } {
        # Delete a worksheet.
        #
        # workbookId  - Identifier of the workbook containing the worksheet.
        # worksheetId - Identifier of the worksheet to delete.
        #
        # No return value.
        # If the number of worksheets before deletion is 1, an error is thrown.
        #
        # See also: DeleteWorksheetByIndex GetWorksheetIdByIndex AddWorksheet

        set count [$workbookId -with { Worksheets } Count]

        if { $count == 1 } {
            error "DeleteWorksheet: Cannot delete last worksheet."
        }

        # Delete the specified worksheet.
        # This will cause alert dialogs to be displayed unless
        # they are turned off.
        set appId [::Cawt::GetApplicationId $workbookId]
        ::Cawt::ShowAlerts $appId false
        $worksheetId Delete
        # Turn the alerts back on.
        ::Cawt::ShowAlerts $appId true
    }

    proc DeleteWorksheetByIndex { workbookId index } {
        # Delete a worksheet identified by it's index.
        #
        # workbookId - Identifier of the workbook containing the worksheet.
        # index      - Index of the worksheet to delete.
        #
        # No return value.
        #
        # The left-most worksheet has index 1.
        # If the index is out of bounds, or the number of worksheets before deletion is 1,
        # an error is thrown.
        #
        # See also: GetNumWorksheets GetWorksheetIdByIndex AddWorksheet

        set count [::Excel::GetNumWorksheets $workbookId]

        if { $count == 1 } {
            error "DeleteWorksheetByIndex: Cannot delete last worksheet."
        }
        if { $index < 1 || $index > $count } {
            error "DeleteWorksheetByIndex: Invalid index $index given."
        }
        # Delete the specified worksheet.
        # This will cause alert dialogs to be displayed unless
        # they are turned off.
        set appId [::Cawt::GetApplicationId $workbookId]
        ::Cawt::ShowAlerts $appId false
        set worksheetId [$workbookId -with { Worksheets } Item [expr $index]]
        $worksheetId Delete
        # Turn the alerts back on.
        ::Cawt::ShowAlerts $appId true
        ::Cawt::Destroy $worksheetId
    }

    proc CopyWorksheet { fromWorksheetId toWorksheetId } {
        # Copy the contents of a worksheet into another worksheet.
        #
        # fromWorksheetId - Identifier of the source worksheet.
        # toWorksheetId   - Identifier of the destination worksheet.
        #
        # Note, that the contents of worksheet toWorksheetId are overwritten.
        #
        # No return value.
        #
        # See also: SelectAll CopyWorksheetBefore CopyWorksheetAfter AddWorksheet

        $fromWorksheetId Activate
        set rangeId [::Excel::SelectAll $fromWorksheetId]
        $rangeId Copy

        $toWorksheetId Activate
        $toWorksheetId Paste

        ::Cawt::Destroy $rangeId
    }

    proc CopyWorksheetBefore { fromWorksheetId beforeWorksheetId { worksheetName "" } } {
        # Copy the contents of a worksheet before another worksheet.
        #
        # fromWorksheetId   - Identifier of the source worksheet.
        # beforeWorksheetId - Identifier of the destination worksheet.
        # worksheetName     - Name of the new worksheet. If no name is specified,
        #                     or an empty string, the naming is done by Excel.
        #
        # Instead of using the identifier of afterWorksheetId, it is also possible
        # to use the numeric index or the special word "end" for specifying the
        # last worksheet.
        #
        # Note, that a new worksheet is generated before worksheet beforeWorksheetId,
        # while CopyWorksheet overwrites the contents of an existing worksheet.
        # The new worksheet is set as the active sheet.
        #
        # Return the identifier of the new worksheet.
        #
        # See also: SelectAll CopyWorksheet CopyWorksheetAfter AddWorksheet

        set fromWorkbookId   [$fromWorksheetId   Parent]
        set beforeWorkbookId [$beforeWorksheetId Parent]

        if { $beforeWorksheetId eq "end" } {
            set beforeWorksheetId [::Excel::GetWorksheetIdByIndex $fromWorkbookId "end"]
        } elseif { [string is integer $beforeWorksheetId] } {
            set index [expr int($beforeWorksheetId)]
            set beforeWorksheetId [::Excel::GetWorksheetIdByIndex $fromWorkbookId ]
        }

        $fromWorksheetId -callnamedargs Copy Before $beforeWorksheetId

        set beforeName [::Excel::GetWorksheetName $beforeWorksheetId]
        set beforeWorksheetIndex [::Excel::GetWorksheetIndexByName $beforeWorkbookId $beforeName]
        set newWorksheetIndex [expr { $beforeWorksheetIndex - 1 }]
        set newWorksheetId [::Excel::GetWorksheetIdByIndex $beforeWorkbookId $newWorksheetIndex]

        if { $worksheetName ne "" } {
            ::Excel::SetWorksheetName $newWorksheetId $worksheetName
        }
        $newWorksheetId Activate
        return $newWorksheetId
    }

    proc CopyWorksheetAfter { fromWorksheetId afterWorksheetId { worksheetName "" } } {
        # Copy the contents of a worksheet after another worksheet.
        #
        # fromWorksheetId  - Identifier of the source worksheet.
        # afterWorksheetId - Identifier of the destination worksheet.
        # worksheetName    - Name of the new worksheet. If no name is specified,
        #                    or an empty string, the naming is done by Excel.
        #
        # Instead of using the identifier of afterWorksheetId, it is also possible
        # to use the numeric index or the special word "end" for specifying the
        # last worksheet.
        #
        # Note, that a new worksheet is generated after worksheet afterWorksheetId,
        # while CopyWorksheet overwrites the contents of an existing worksheet.
        # The new worksheet is set as the active sheet.
        #
        # Return the identifier of the new worksheet.
        #
        # See also: SelectAll CopyWorksheet CopyWorksheetBefore AddWorksheet

        set fromWorkbookId  [$fromWorksheetId  Parent]
        set afterWorkbookId [$afterWorksheetId Parent]

        if { $afterWorksheetId eq "end" } {
            set afterWorksheetId [::Excel::GetWorksheetIdByIndex $fromWorkbookId "end"]
        } elseif { [string is integer $afterWorksheetId] } {
            set index [expr int($afterWorksheetId)]
            set afterWorksheetId [::Excel::GetWorksheetIdByIndex $fromWorkbookId ]
        }

        $fromWorksheetId -callnamedargs Copy After $afterWorksheetId

        set afterName [::Excel::GetWorksheetName $afterWorksheetId]
        set afterWorksheetIndex [::Excel::GetWorksheetIndexByName $afterWorkbookId $afterName]
        set newWorksheetIndex [expr { $afterWorksheetIndex + 1 }]
        set newWorksheetId [::Excel::GetWorksheetIdByIndex $afterWorkbookId $newWorksheetIndex]

        if { $worksheetName ne "" } {
            ::Excel::SetWorksheetName $newWorksheetId $worksheetName
        }
        $newWorksheetId Activate
        return $newWorksheetId
    }

    proc GetWorksheetIdByIndex { workbookId index { activate true } } {
        # Find a worksheet by it's index.
        #
        # workbookId - Identifier of the workbook containing the worksheet.
        # index      - Index of the worksheet to find.
        # activate   - true: Activate the found worksheet.
        #              false: Just return the identifier.
        #
        # Return the identifier of the found worksheet.
        # The left-most worksheet has index 1.
        # Instead of using the numeric index the special word "end" may
        # be used to specify the last worksheet.
        # If the index is out of bounds an error is thrown.
        #
        # See also: GetNumWorksheets GetWorksheetIdByName AddWorksheet

        set count [::Excel::GetNumWorksheets $workbookId]
        if { $index eq "end" } {
            set index $count
        } else {
            if { $index < 1 || $index > $count } {
                error "GetWorksheetIdByIndex: Invalid index $index given."
            }
        }
        set worksheetId [$workbookId -with { Worksheets } Item [expr $index]]
        if { $activate } {
            $worksheetId Activate
        }
        return $worksheetId
    }

    proc GetWorksheetIdByName { workbookId worksheetName { activate true } } {
        # Find a worksheet by it's name.
        #
        # workbookId    - Identifier of the workbook containing the worksheet.
        # worksheetName - Name of the worksheet to find.
        # activate      - true: Activate the found worksheet.
        #                 false: Just return the identifier.
        #
        # Return the identifier of the found worksheet.
        # If a worksheet with given name does not exist an error is thrown.
        #
        # See also: GetNumWorksheets GetWorksheetIndexByName GetWorksheetIdByIndex AddWorksheet

        set worksheets [$workbookId Worksheets]
        set count [$worksheets Count]

        for { set i 1 } { $i <= $count } { incr i } {
            set worksheetId [$worksheets Item [expr $i]]
            if { $worksheetName eq [$worksheetId Name] } {
                ::Cawt::Destroy $worksheets
                if { $activate } {
                    $worksheetId Activate
                }
                return $worksheetId
            }
            ::Cawt::Destroy $worksheetId
        }
        error "GetWorksheetIdByName: No worksheet with name $worksheetName"
    }

    proc GetWorksheetIndexByName { workbookId worksheetName { activate true } } {
        # Find a worksheet index by it's name.
        #
        # workbookId    - Identifier of the workbook containing the worksheet.
        # worksheetName - Name of the worksheet to find.
        # activate      - true: Activate the found worksheet.
        #                 false: Just return the index.
        #
        # Return the index of the found worksheet.
        # The left-most worksheet has index 1.
        # If a worksheet with given name does not exist an error is thrown.
        #
        # See also: GetNumWorksheets GetWorksheetIdByIndex GetWorksheetIdByName AddWorksheet

        set worksheets [$workbookId Worksheets]
        set count [$worksheets Count]

        for { set i 1 } { $i <= $count } { incr i } {
            set worksheetId [$worksheets Item [expr $i]]
            if { $worksheetName eq [$worksheetId Name] } {
                ::Cawt::Destroy $worksheets
                if { $activate } {
                    $worksheetId Activate
                }
                return $i
            }
            ::Cawt::Destroy $worksheetId
        }
        error "GetWorksheetIdByName: No worksheet with name $worksheetName"
    }

    proc SetWorksheetName { worksheetId name } {
        # Set the name of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        # name        - Name of the worksheet.
        #
        # No return value.
        #
        # See also: GetWorksheetName AddWorksheet

        $worksheetId Name $name
    }

    proc GetWorksheetName { worksheetId } {
        # Return the name of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        #
        # See also: SetWorksheetName AddWorksheet

        return [$worksheetId Name]
    }

    proc IsWorksheetProtected { worksheetId } {
        # Check, if a worksheet is content protected.
        #
        # worksheetId - Identifier of the worksheet to be checked.
        #
        # Return true, if the worksheet is protected, otherwise false.
        #
        # See also: AddWorksheet

        if { [$worksheetId ProtectContents] } {
            return true
        } else {
            return false
        }
    }

    proc IsWorksheetVisible { worksheetId } {
        # Check, if a worksheet is visible.
        #
        # worksheetId - Identifier of the worksheet to be checked.
        #
        # Return true, if the worksheet is visible, otherwise false.
        #
        # See also: AddWorksheet

        if { [$worksheetId Visible] == $::Excel::xlSheetVisible } {
            return true
        } else {
            return false
        }
    }

    proc GetNumWorksheets { workbookId } {
        # Return the number of worksheets in a workbook.
        #
        # workbookId - Identifier of the workbook.
        #
        # See also: AddWorksheet OpenWorkbook

        return [$workbookId -with { Worksheets } Count]
    }

    proc SetWorksheetOrientation { worksheetId orientation } {
        # Set the orientation of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        # orientation - Value of enumeration type XlPageOrientation (see excelConst.tcl).
        #               Possible values: xlLandscape or xlPortrait.
        #
        # No return value.
        #
        # See also: AddWorksheet

        $worksheetId -with { PageSetup } Orientation $orientation
    }

    proc SetWorksheetFitToPages { worksheetId { wide 1 } { tall 1 } } {
        # Adjust a worksheet to fit onto given number of pages.
        #
        # worksheetId - Identifier of the worksheet.
        # wide        - The number of pages in horizontal direction.
        # tall        - The number of pages in vertical direction.
        #
        # When using the default values for wide and tall, the worksheet is adjusted
        # to fit onto exactly one piece of paper.
        #
        # No return value.
        #
        # See also: AddWorksheet

        set pageSetup [$worksheetId PageSetup]
        $pageSetup Zoom [::Cawt::TclBool false]
        $pageSetup FitToPagesWide $wide
        $pageSetup FitToPagesTall $tall
        ::Cawt::Destroy $pageSetup
    }

    proc SetWorksheetZoom { worksheetId { zoom 100 } } {
        # Set the zoom factor for printing of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        # zoom        - The zoom factor in percent as an integer value.
        #
        # Valid zoom values are in the range [10, 400].
        #
        # No return value.
        #
        # See also: AddWorksheet

        $worksheetId -with { PageSetup } Zoom [expr int($zoom)]
    }

    proc SetWorksheetTabColor { worksheetId r g b } {
        # Set the color of the tab of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        # r           - Red component of the tab color.
        # g           - Green component of the tab color.
        # b           - Blue component of the tab color.
        #
        # The r, g and b values are specified as integers in the
        # range [0, 255].
        #
        # No return value.
        #
        # See also: SetRangeTextColor ::Cawt::RgbToColor GetWorksheetIdByIndex

        set color [::Cawt::RgbToColor $r $g $b]
        $worksheetId -with { Tab } Color $color
    }

    proc UnhideWorksheet { worksheetId { r 0 } { g 128 } { b 0 } } {
        # Unhide a worksheet, if it is hidden.
        #
        # worksheetId - Identifier of the worksheet.
        # r           - Red component of the tab color.
        # g           - Green component of the tab color.
        # b           - Blue component of the tab color.
        #
        # If the worksheet is hidden, it is made visible and the tab color is set
        # to the specified color.
        #
        # The r, g and b values are specified as integers in the
        # range [0, 255].
        #
        # No return value.
        #
        # See also: SetWorksheetTabColor IsWorksheetVisible ::Cawt::RgbToColor

        if { ! [::Excel::IsWorksheetVisible $worksheetId] } {
            if { [$worksheetId -with { Parent } ProtectStructure] } {
                error "Unable to unhide because the Workbook's structure is protected."
            } else {
                $worksheetId Visible $::Excel::xlSheetVisible
                ::Excel::SetWorksheetTabColor $worksheetId $r $g $b
            }
        }
    }

    proc GetCellsId { worksheetId } {
        # Return the cells identifier of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        #
        # Return the range of all cells from a worksheet. This corresponds to the
        # method Cells() of the Worksheet object.

        set cellsId [$worksheetId Cells]
        return $cellsId
    }

    proc GetCellIdByIndex { worksheetId row col } {
        # Return a cell of a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        # row         - Row number. Row numbering starts with 1.
        # col         - Column number. Column numbering starts with 1.
        #
        # Return the cell identifier of the cell with index (row, col).
        #
        # See also: SelectCellByIndex AddWorksheet

        set cellsId [::Excel::GetCellsId $worksheetId]
        set cell [$cellsId Item [expr {int($row)}] [expr {int($col)}]]
        ::Cawt::Destroy $cellsId
        return $cell
    }

    proc SelectCellByIndex { worksheetId row col { visSel false } } {
        # Select a cell by it's row/column index.
        #
        # worksheetId - Identifier of the worksheet.
        # row         - Row number. Row numbering starts with 1.
        # col         - Column number. Column numbering starts with 1.
        # visSel      - true: See the selection in the user interface.
        #               false: The selection is not visible.
        #
        # Return the identifier of the cell as a range identifier.
        #
        # See also: SelectRangeByIndex AddWorksheet

        return [::Excel::SelectRangeByIndex $worksheetId $row $col $row $col $visSel]
    }

    proc ShowCellByIndex { worksheetId row col } {
        # Show a cell identified by it's row/column index.
        #
        # worksheetId - Identifier of the worksheet.
        # row         - Row number. Row numbering starts with 1.
        # col         - Column number. Column numbering starts with 1.
        #
        # Set the scrolling, so that the cell is show at the upper left corner.
        #
        # See also: SelectCellByIndex

        if { $row <= 0 } {
            error "Row number $row is invalid."
        }
        if { $col <= 0 } {
            error "Column number $col is invalid."
        }
        $worksheetId Activate
        set actWin [[::Cawt::GetApplicationId $worksheetId] ActiveWindow]
        $actWin ScrollColumn $col
        $actWin ScrollRow $row
    }

    proc SetHyperlink { worksheetId row col link { textDisplay "" } } {
        # Insert a hyperlink into a worksheet.
        #
        # worksheetId - Identifier of the worksheet where the hyperlink is inserted.
        # row         - Row number. Row numbering starts with 1.
        # col         - Column number. Column numbering starts with 1.
        # link        - URL of the hyperlink.
        # textDisplay - Text to be displayed instead of the URL.
        #
        # URL's are specified as strings. "file://myLinkedFile" specifies a link
        # to a local file.
        #
        # No return value.
        #
        # See also: AddWorksheet

        variable excelVersion

        if { $textDisplay eq "" } {
            set textDisplay $link
        }

        set rangeId [SelectRangeByIndex $worksheetId $row $col $row $col]
        set hyperId [$worksheetId Hyperlinks]

        # Add(Anchor As Object, Address As String, [SubAddress],
        # [ScreenTip], [TextToDisplay]) As Object
        if { $excelVersion eq "8.0" } {
            $hyperId -callnamedargs Add \
                     Anchor $rangeId \
                     Address $link
        } else {
            $hyperId -callnamedargs Add \
                     Anchor $rangeId \
                     Address $link \
                     TextToDisplay $textDisplay
        }
        ::Cawt::Destroy $hyperId
        ::Cawt::Destroy $rangeId
    }

    proc InsertImage { worksheetId imgFileName { row 1 } { col 1 } } {
        # Insert an image into a worksheet.
        #
        # worksheetId - Identifier of the worksheet where the image is inserted.
        # imgFileName - File name of the image (as absolute path).
        # row         - Row number. Row numbering starts with 1.
        # col         - Column number. Column numbering starts with 1.
        #
        # The file name of the image must be an absolute pathname. Use a
        # construct like [file join [pwd] "myImage.gif"] to insert
        # images from the current directory.
        #
        # Return the identifier of the inserted image.
        #
        # See also: ScaleImage

        set cellId [SelectCellByIndex $worksheetId $row $col true]
        set pictures [$worksheetId Pictures]
        set fileName [file nativename $imgFileName]
        set picId [$pictures Insert $fileName]
        ::Cawt::Destroy $cellId
        ::Cawt::Destroy $pictures
        return $picId
    }

    proc ScaleImage { picId scaleWidth scaleHeight } {
        # Scale an image.
        #
        # picId       - Identifier of the image.
        # scaleWidth  - Horizontal scale factor.
        # scaleHeight - Vertical scale factor.
        #
        # The scale factors are floating point values. 1.0 means no scaling.
        #
        # No return value.
        #
        # See also: InsertImage

        set rangeId [$picId ShapeRange]
        $rangeId ScaleWidth  [expr double($scaleWidth)]  [::Cawt::TclInt true]
        $rangeId ScaleHeight [expr double($scaleHeight)] [::Cawt::TclInt true]
        ::Cawt::Destroy $rangeId
    }

    proc SetCellValue { worksheetId row col val { fmt "text" } { subFmt "" } } {
        # Set the value of a cell.
        #
        # worksheetId - Identifier of the worksheet.
        # row         - Row number. Row numbering starts with 1.
        # col         - Column number. Column numbering starts with 1.
        # val         - String value of the cell.
        # fmt         - Format of the cell. Possible values: "text", "int", "real".
        # subFmt      - Formatting option of the floating-point value (see SetRangeFormat).
        #
        # The value to be inserted is interpreted either as string, integer or
        # floating-point number according to the formats specified in "fmt" and "subFmt".
        #
        # See also: GetCellValue SetRowValues SetMatrixValues

        set cellsId [::Excel::GetCellsId $worksheetId]
        set cellId [::Excel::GetCellIdByIndex $worksheetId $row $col]
        SetRangeFormat $cellId $fmt $subFmt
        if { $fmt eq "text" } {
            $cellsId Item [expr {int($row)}] [expr {int($col)}] [format "%s" $val]
        } elseif { $fmt eq "int" } {
            $cellsId Item [expr {int($row)}] [expr {int($col)}] [expr {int($val)}]
        } elseif { $fmt eq "real" } {
            $cellsId Item [expr {int($row)}] [expr {int($col)}] [expr {double($val)}]
        } else {
            error "SetCellValue: Unknown format $fmt"
        }
        ::Cawt::Destroy $cellId
        ::Cawt::Destroy $cellsId
    }

    proc GetCellValue { worksheetId row col { fmt "text" } } {
        # Return the value of a cell.
        #
        # worksheetId - Identifier of the worksheet.
        # row         - Row number. Row numbering starts with 1.
        # col         - Column number. Column numbering starts with 1.
        # fmt         - Format of the cell. Possible values: "text", "int", "real".
        #
        # Depending on the format the value of the cell is returned as a string, integer number
        # or a floating-point number.
        # If the value could not be retrieved, an error is thrown.
        #
        # See also: SetCellValue ColumnCharToInt

        set cellsId [::Excel::GetCellsId $worksheetId]
        set cell [$cellsId Item [expr {int($row)}] [expr {int($col)}]]
        set retVal [catch {$cell Value} val]
        if { $retVal != 0 } {
            error "GetCellValue: Unable to get value of cell ($row, $col)"
        }
        ::Cawt::Destroy $cell
        ::Cawt::Destroy $cellsId
        if { $fmt eq "text" } {
            return $val
        } elseif { $fmt eq "int" } {
            return [expr {int ($val)}]
        } elseif { $fmt eq "real" } {
            return [expr {double ($val)}]
        } else {
            error "GetCellValue: Unknown format $fmt"
        }
    }

    proc SetRowValues { worksheetId row valList { startCol 1 } { numVals 0 } } {
        # Insert row values from a Tcl list.
        #
        # worksheetId - Identifier of the worksheet.
        # row         - Row number. Row numbering starts with 1.
        # valList     - List of values to be inserted.
        # startCol    - Column number of insertion start. Column numbering starts with 1.
        # numVals     - Negative or zero: All list values are inserted.
        #               Positive: numVals columns are filled with the list values
        #               (starting at list index 0).
        #
        # No return value. If valList is an empty list, an error is thrown.
        #
        # See also: GetRowValues SetColumnValues SetCellValue ColumnCharToInt

        set len [llength $valList]
        if { $numVals > 0 } {
            if { $numVals < $len } {
                set len $numVals
            }
        }

        set cellId [::Excel::SelectRangeByIndex $worksheetId \
                    $row $startCol $row [expr {$startCol + $len -1}]]
        $cellId Value2 $valList
        ::Cawt::Destroy $cellId
    }

    proc GetRowValues { worksheetId row { startCol 0 } { numVals 0 } } {
        # Return row values as a Tcl list.
        #
        # worksheetId - Identifier of the worksheet.
        # row         - Row number. Row numbering starts with 1.
        # startCol    - Column number of start. Column numbering starts with 1.
        #               Negative or zero: Start at first available column.
        # numVals     - Negative or zero: All available row values are returned.
        #               Positive: Only numVals values of the row are returned.
        #
        # Note, that the functionality of this procedure has changed slightly with
        # CAWT versions greater than 1.0.5:
        # If "startCol" is not specified, "startCol" is not set to 1, but it is set to
        # the first available row.
        # Possible incompatibility.
        #
        # Return the values of the specified row or row range as a Tcl list.
        #
        # See also: SetRowValues GetColumnValues GetCellValue ColumnCharToInt GetFirstUsedColumn

        if { $startCol <= 0 } {
            set startCol [::Excel::GetFirstUsedColumn $worksheetId]
        }
        if { $numVals <= 0 } {
            set numVals [expr { $startCol + [::Excel::GetLastUsedColumn $worksheetId] - 1 }]
        }
        set valList [list]
        set col $startCol
        set ind 0
        while { $ind < $numVals } {
            lappend valList [::Excel::GetCellValue $worksheetId $row $col]
            incr ind
            incr col
        }
        return $valList
    }

    proc SetColumnWidth { worksheetId col { width 0 } } {
        # Set the width of a column.
        #
        # worksheetId - Identifier of the worksheet.
        # col         - Column number. Column numbering starts with 1.
        # width       - A positive value specifies the column's width in average-size characters
        #               of the widget's font. A value of zero specifies that the column's width
        #               fits automatically the width of all elements in the column.
        #
        # No return value.
        #
        # See also: SetColumnsWidth ColumnCharToInt

        set cell [SelectRangeByIndex $worksheetId 1 $col 1 $col]
        set curCol [$cell EntireColumn]
        if { $width == 0 } {
            [$curCol Columns] AutoFit
        } else {
            $curCol ColumnWidth $width
        }
        ::Cawt::Destroy $curCol
        ::Cawt::Destroy $cell
    }

    proc SetColumnsWidth { worksheetId startCol endCol { width 0 } } {
        # Set the width of a range of columns.
        #
        # worksheetId - Identifier of the worksheet.
        # startCol    - Range start column number. Column numbering starts with 1.
        # endCol      - Range end column number. Column numbering starts with 1.
        # width       - A positive value specifies the column's width in average-size characters
        #               of the widget's font. A value of zero specifies that the column's width
        #               fits automatically the width of all elements in the column.
        #
        # No return value.
        #
        # See also: SetColumnWidth ColumnCharToInt

        for { set c $startCol } { $c <= $endCol } { incr c } {
            SetColumnWidth $worksheetId $c $width
        }
    }

    proc SetColumnValues { worksheetId col valList { startRow 1 } { numVals 0 } } {
        # Insert column values from a Tcl list.
        #
        # worksheetId - Identifier of the worksheet.
        # col         - Column number. Column numbering starts with 1.
        # valList     - List of values to be inserted.
        # startRow    - Row number of insertion start. Row numbering starts with 1.
        # numVals     - Negative or zero: All list values are inserted.
        #               Positive: numVals rows are filled with the list values
        #               (starting at list index 0).
        #
        # No return value.
        #
        # See also: GetColumnValues SetRowValues SetCellValue ColumnCharToInt

        set len [llength $valList]
        if { $numVals > 0 } {
            if { $numVals < $len } {
                set len $numVals
            }
        }

        for { set i 0 } { $i < $len } { incr i } {
            lappend valListList [list [lindex $valList $i]]
        }
        set cellId [::Excel::SelectRangeByIndex $worksheetId \
                    $startRow $col [expr {$startRow + $len - 1}] $col]
        $cellId Value2 $valListList
        ::Cawt::Destroy $cellId
    }

    proc GetColumnValues { worksheetId col { startRow 0 } { numVals 0 } } {
        # Return column values as a Tcl list.
        #
        # worksheetId - Identifier of the worksheet.
        # col         - Column number. Column numbering starts with 1.
        # startRow    - Row number of start. Row numbering starts with 1.
        #               Negative or zero: Start at first available row.
        # numVals     - Negative or zero: All available column values are returned.
        #               Positive: Only numVals values of the column are returned.
        #
        # Note, that the functionality of this procedure has changed slightly with
        # CAWT versions greater than 1.0.5:
        # If "startRow" is not specified, "startRow" is not set to 1, but it is set to
        # the first available row.
        # Possible incompatibility.
        #
        # Return the values of the specified column or column range as a Tcl list.
        #
        # See also: SetColumnValues GetRowValues GetCellValue ColumnCharToInt GetFirstUsedRow

        if { $startRow <= 0 } {
            set startRow [::Excel::GetFirstUsedRow $worksheetId]
        }
        if { $numVals <= 0 } {
            set numVals [expr { $startRow + [::Excel::GetLastUsedRow $worksheetId] - 1 }]
        }
        set valList [list]
        set row $startRow
        set ind 0
        while { $ind < $numVals } {
            lappend valList [GetCellValue $worksheetId $row $col]
            incr ind
            incr row
        }
        return $valList
    }

    proc SetMatrixValues { worksheetId matrixList { startRow 1 } { startCol 1 } } {
        # Insert matrix values into a worksheet.
        #
        # worksheetId - Identifier of the worksheet.
        # matrixList  - Matrix with table data.
        # startRow    - Row number of insertion start. Row numbering starts with 1.
        # startCol    - Column number of insertion start. Column numbering starts with 1.
        #
        # The matrix data must be stored as a list of lists. Each sub-list contains
        # the values for the row values.
        # The main (outer) list contains the rows of the matrix.
        # Example:
        # { { R1_C1 R1_C2 R1_C3 } { R2_C1 R2_C2 R2_C3 } }
        #
        # No return value.
        #
        # See also: GetMatrixValues SetRowValues SetColumnValues

        set numCols [llength [lindex $matrixList 0]]
        set numRows [llength $matrixList]

        set cellId [::Excel::SelectRangeByIndex $worksheetId \
                    $startRow $startCol \
                    [expr {$startRow + $numRows -1}] [expr {$startCol + $numCols -1}]]
        $cellId Value2 $matrixList
        ::Cawt::Destroy $cellId
    }

    proc GetMatrixValues { worksheetId row1 col1 row2 col2 } {
        # Return worksheet table values as a matrix.
        #
        # worksheetId - Identifier of the worksheet.
        # row1        - Row number of upper-left corner of the cell range.
        # col1        - Column number of upper-left corner of the cell range.
        # row2        - Row number of lower-right corner of the cell range.
        # col2        - Column number of lower-right corner of the cell range.
        #
        # See also: SetMatrixValues GetRowValues GetColumnValues

        set cellId [::Excel::SelectRangeByIndex $worksheetId \
                    $row1 $col1 $row2 $col2 true]
        set matrixList [$cellId Value2]

        ::Cawt::Destroy $cellId
        return $matrixList
    }
}
