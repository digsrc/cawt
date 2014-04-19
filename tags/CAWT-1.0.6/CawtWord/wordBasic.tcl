# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval ::Word {

    variable wordVersion "0.0"
    variable wordAppName "Word.Application"
    variable _ruffdoc

    lappend _ruffdoc Introduction {
        The Word namespace provides commands to control Microsoft Word.
    }

    proc TrimString { str } {
        # Trim a string.
        #
        # str - String to be trimmed.
        #
        # The string is trimmed from the left and right side.
        # Trimmed characters are: Whitespaces, BEL (0x7) and CR (0xD).
        #
        # Return the trimmed string.

        set str [string trim $str]
        set str [string trim $str [format "%c" 0x7]]
        set str [string trim $str [format "%c" 0xD]]
        return $str
    }

    proc FindString { rangeId str { matchCase true } } {
        # Find a string in a text range.
        #
        # rangeId   - Identifier of the text range.
        # str       - Search string.
        # matchCase - Flag indicating case sensitive search.
        #
        # Return true, if string was found. Otherwise false.
        # If the string was found, the selection is set to the found string.
        #
        # See also: ReplaceString GetSelectionRange

        set myFind [$rangeId Find]
        # Execute([FindText], [MatchCase], [MatchWholeWord], [MatchWildcards],
        # [MatchSoundsLike], [MatchAllWordForms], [Forward], [Wrap], [Format],
        # [ReplaceWith], [Replace], [MatchKashida], [MatchDiacritics],
        # [MatchAlefHamza], [MatchControl]) As Boolean
        set retVal [$myFind -callnamedargs Execute \
                            FindText $str \
                            MatchCase [::Cawt::TclBool $matchCase] \
                            Wrap $::Word::wdFindStop \
                            Forward True]
        ::Cawt::Destroy $myFind
        if { $retVal } {
            return true
        }
        return false
    }

    proc ReplaceString { rangeId searchStr replaceStr \
                        { howMuch "one" } { matchCase true } } {
        # Replace a string in a text range. Simple case.
        #
        # rangeId    - Identifier of the text range.
        # searchStr  - Search string.
        # replaceStr - Replacement string.
        # howMuch    - "one" to replace first occurence only. "all" to replace all occurences.
        # matchCase  - Flag indicating case sensitive search.
        #
        # Return true, if string could be replaced, i.e. the search string was found.
        # Otherwise false.
        #
        # See also: SearchString ReplaceByProc

        set howMuchEnum $::Word::wdReplaceOne
        if { $howMuch ne "one" } {
            set howMuchEnum $::Word::wdReplaceAll
        }
        set myFind [$rangeId Find]
        # See proc FindString for parameter list of Execute command.
        set retVal [$myFind -callnamedargs Execute \
                            FindText $searchStr \
                            ReplaceWith $replaceStr \
                            Replace $howMuchEnum \
                            Wrap $::Word::wdFindStop \
                            MatchCase [::Cawt::TclBool $matchCase] \
                            Forward True]
        ::Cawt::Destroy $myFind
        if { $retVal } {
            return true
        }
        return false
    }

    proc ReplaceByProc { rangeId str func args } {
        # Replace a string in a text range. Generic case.
        #
        # rangeId - Identifier of the text range.
        # str     - Search string.
        # func    - Replacement procedure.
        # args    - Arguments for replacement procedure.
        #
        # Search for string "str" in the range "rangeId". For each
        # occurence found, call procedure "func" with the range of
        # the found occurence and additional parameters specified in
        # "args". The procedures which can be used for "func" must
        # therefore have the following signature:
        # proc SetRangeXYZ rangeId param1 param2 ...
        # See example Word-04-Find.tcl for an example.
        #
        # No return value.
        #
        # See also: SearchString ReplaceString

        set myFind [$rangeId Find]
        set count 0
        while { 1 } {
            # See proc FindString for parameter list of Execute command.
            set retVal [$myFind -callnamedargs Execute \
                                FindText $str \
                                MatchCase True \
                                Forward True]
            if { ! $retVal } {
                break
            }
            eval $func $rangeId $args
            incr count
        }
        ::Cawt::Destroy $myFind
    }

    proc GetSelectionRange { docId } {
        # Return the text range representing the current selection.
        #
        # docId - Identifier of the document.
        #
        # See also: GetStartRange GetEndRange

        return [$docId -with { ActiveWindow } Selection]
    }

    proc GetStartRange { docId } {
        # Return the text range representing the start of the document.
        #
        # docId - Identifier of the document.
        #
        # See also: GetSelectionRange GetEndRange

        set rangeId [::Word::GetSelectionRange $docId]
        $rangeId Start 0
        $rangeId End 0
        return $rangeId
    }

    proc GetEndRange { docId } {
        # Return the text range representing the end of the document.
        #
        # docId - Identifier of the document.
        #
        # Note: This corresponds to the built-in bookmark \endofdoc.
        #
        # See also: GetSelectionRange GetStartRange

        set bookMarks [$docId Bookmarks]
        set endOfDoc  [$bookMarks Item "\\endofdoc"]
        set endRange  [$endOfDoc Range]
        ::Cawt::Destroy $endOfDoc
        ::Cawt::Destroy $bookMarks
        return $endRange
    }

    proc PrintRange { rangeId { msg "Range: " } } {
        # Print the indices of a text range.
        #
        # rangeId - Identifier of the text range.
        # msg     - String printed in front of the indices.
        #
        # The range identifiers are printed onto standard output.
        #
        # No return value.
        #
        # See also: GetRangeStartIndex GetRangeEndIndex

        puts [format "%s %d %d" $msg \
              [::Word::GetRangeStartIndex $rangeId] [::Word::GetRangeEndIndex $rangeId]]
    }

    proc GetRangeStartIndex { rangeId } {
        # Return the start index of a text range.
        #
        # rangeId - Identifier of the text range.
        #
        # See also: GetRangeEndIndex PrintRange

        return [$rangeId Start]
    }

    proc GetRangeEndIndex { rangeId } {
        # Return the end index of a text range.
        #
        # rangeId - Identifier of the text range.
        #
        # See also: GetRangeStartIndex PrintRange

        return [$rangeId End]
    }

    proc SetRangeStartIndex { docId rangeId index } {
        # Set the start index of a text range.
        #
        # docId   - Identifier of the document.
        # rangeId - Identifier of the text range.
        # index   - Index for the range start.
        #
        # Index is either an integer value or string "begin" to
        # use the start of the document.
        #
        # No return value.
        #
        # See also: SetRangeEndIndex GetRangeStartIndex

        if { $index eq "begin" } {
            set index 0
        }
        $rangeId Start $index
    }

    proc SetRangeEndIndex { docId rangeId index } {
        # Set the end index of a text range.
        #
        # docId   - Identifier of the document.
        # rangeId - Identifier of the text range.
        # index   - Index for the range end.
        #
        # Index is either an integer value or string "end" to
        # use the end of the document.
        #
        # No return value.
        #
        # See also: SetRangeBeginIndex GetRangeEndIndex

        if { $index eq "end" } {
            set index [[GetEndRange $docId] End]
        }
        $rangeId End $index
    }

    proc ExtendRange { docId rangeId { startIncr 0 } { endIncr 0 } } {
        # Extend the range indices of a text range.
        #
        # docId     - Identifier of the document.
        # rangeId   - Identifier of the text range.
        # startIncr - Increment of the range start index.
        # endIncr   - Increment of the range end index.
        #
        # Increment is either an integer value or strings "begin" or "end" to
        # use the start or end of the document.
        #
        # Return the new extended range.
        #
        # See also: SetRangeBeginIndex SetRangeEndIndex

        set startIndex [::Word::GetRangeStartIndex $rangeId]
        set endIndex   [::Word::GetRangeEndIndex   $rangeId]
        if { [string is integer $startIncr] } {
            set startIndex [expr $startIndex + $startIncr]
        } elseif { $startIncr eq "begin" } {
            set startIndex 0
        }
        if { [string is integer $endIncr] } {
            set endIndex [expr $endIndex + $endIncr]
        } elseif { $endIncr eq "end" } {
            set endIndex [[GetEndRange $docId] End]
        }
        $rangeId Start $startIndex
        $rangeId End $endIndex
        return $rangeId
    }

    proc SetRangeFontBold { rangeId { onOff true } } {
        # Toggle the bold font style of a text range.
        #
        # rangeId - Identifier of the text range.
        # onOff   - true: Set bold style on.
        #           false: Set bold style off.
        #
        # No return value.
        #
        # See also: SetRangeFontItalic

        $rangeId -with { Font } Bold [::Cawt::TclInt $onOff]
    }

    proc SetRangeFontItalic { rangeId { onOff true } } {
        # Toggle the italic font style of a text range.
        #
        # rangeId - Identifier of the text range.
        # onOff   - true: Set italic style on.
        #           false: Set italic style off.
        #
        # No return value.
        #
        # See also: SetRangeFontBold

        $rangeId -with { Font } Italic [::Cawt::TclInt $onOff]
    }

    proc SetRangeHorizontalAlignment { rangeId align } {
        # Set the horizontal alignment of a text range.
        #
        # rangeId - Identifier of the text range.
        # align   - Value of enumeration type WdParagraphAlignment (see wordConst.tcl).
        #
        # No return value.
        #
        # See also: SetRangeHighlightColorByEnum

        $rangeId -with { ParagraphFormat } Alignment $align
    }

    proc SetRangeHighlightColorByEnum { rangeId colorEnum } {
        # Set the highlight color of a text range.
        #
        # rangeId   - Identifier of the text range.
        # colorEnum - Value of enumeration type WdColor (see wordConst.tcl).
        #
        # No return value.
        #
        # See also: SetRangeBackgroundColorByEnum

        $rangeId HighlightColorIndex $colorEnum
    }

    proc SetRangeBackgroundColorByEnum { rangeId colorEnum } {
        # Set the background color of a text range.
        #
        # rangeId   - Identifier of the text range.
        # colorEnum - Value of enumeration type WdColor (see wordConst.tcl).
        #
        # No return value.
        #
        # See also: SetRangeBackgroundColor SetRangeHighlightColorByEnum

        $rangeId -with { Cells Shading } BackgroundPatternColor $colorEnum
    }

    proc SetRangeBackgroundColor { rangeId r g b } {
        # Set the background color of a text range.
        #
        # rangeId - Identifier of the text range.
        # r       - Red component of the text color.
        # g       - Green component of the text color.
        # b       - Blue component of the text color.
        #
        # The r, g and b values are specified as integers in the
        # range [0, 255].
        #
        # No return value.
        #
        # See also: SetRangeBackgroundColorByEnum SetRangeHighlightColorByEnum

        $rangeId -with { Cells Shading } BackgroundPatternColor \
                                   [::Cawt::RgbToColor $r $g $b]
    }

    proc GetVersion { appId { useString false } } {
        # Return the version of a Word application.
        #
        # appId     - Identifier of the Word instance.
        # useString - true: Return the version name (ex. "Word 2000").
        #             false: Return the version number (ex. "9.0").
        #
        # Both version name and version number are returned as strings.
        # Version number is in a format, so that it can be evaluated as a
        # floating point number.

        array set map {
            "7.0"  "Word 95"
            "8.0"  "Word 97"
            "9.0"  "Word 2000"
            "10.0" "Word 2002"
            "11.0" "Word 2003"
            "12.0" "Word 2007"
            "14.0" "Word 2010"
            "15.0" "Word 2013"
        }
        set version [$appId Version]
        if { $useString } {
            if { [info exists map($version)] } {
                return $map($version)
            } else {
                return "Unknown Word version"
            }
        } else {
            return $version
        }
    }

    proc GetExtString { appId } {
        # Return the default extension of a Word file.
        #
        # appId - Identifier of the Word instance.
        #
        # Starting with Word 12 (2007) this is the string ".docx".
        # In previous versions it was ".doc".

        # appId is only needed, so we are sure, that wordVersion is initialized.
 
        variable wordVersion

        if { $wordVersion >= 12.0 } {
            return ".docx"
        } else {
            return ".doc"
        }
    }

    proc OpenNew { { visible true } { width -1 } { height -1 } } {
        # Open a new Word instance.
        #
        # visible - true: Show the application window.
        #           false: Hide the application window.
        # width   - Width of the application window. If negative, open with last used width.
        # height  - Height of the application window. If negative, open with last used height.
        #
        # Return the identifier of the new Word application instance.
        #
        # See also: Open Quit Visible

        variable wordAppName
	variable wordVersion

        set appId [::Cawt::GetOrCreateApp $wordAppName false]
        set wordVersion [::Word::GetVersion $appId]
        ::Word::Visible $appId $visible
        if { $width >= 0 } {
            $appId Width [expr $width]
        }
        if { $height >= 0 } {
            $appId Height [expr $height]
        }
        return $appId
    }

    proc Open { { visible true } { width -1 } { height -1 } } {
        # Open a Word instance. Use an already running Word, if available.
        #
        # visible - true: Show the application window.
        #           false: Hide the application window.
        # width   - Width of the application window. If negative, open with last used width.
        # height  - Height of the application window. If negative, open with last used height.
        #
        # Return the identifier of the Word application instance.
        #
        # See also: OpenNew Quit Visible

        variable wordAppName
	variable wordVersion

        set appId [::Cawt::GetOrCreateApp $wordAppName true]
        set wordVersion [::Word::GetVersion $appId]
        ::Word::Visible $appId $visible
        if { $width >= 0 } {
            $appId Width [expr $width]
        }
        if { $height >= 0 } {
            $appId Height [expr $height]
        }
        return $appId
    }

    proc Quit { appId { showAlert true } } {
        # Quit a Word application instance.
        #
        # appId     - Identifier of the Word instance.
        # showAlert - true: Show an alert window, if there are unsaved changes.
        #             false: Quit without saving any changes.
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
        # Toggle the visibilty of a Word application window.
        #
        # appId   - Identifier of the Word instance.
        # visible - true: Show the application window.
        #           false: Hide the application window.
        #
        # No return value.
        #
        # See also: Open OpenNew

        $appId Visible [::Cawt::TclInt $visible]
    }

    proc Close { docId } {
        # Close a document without saving changes.
        #
        # docId - Identifier of the document.
        #
        # Use the SaveAs method before closing, if you want to save changes.
        #
        # No return value.
        #
        # See also: SaveAs

        $docId Close [::Cawt::TclBool false]
    }

    proc UpdateFields { docId } {
        # Update all fields of a document.
        #
        # docId - Identifier of the document.
        #
        # No return value.
        #
        # See also: SaveAs

        set rangeId [::Word::GetStartRange $docId]
        $rangeId WholeStory
        $rangeId -with { Fields } Update
    }

    proc SaveAs { docId fileName { fmt "" } } {
        # Save a document to a Word file.
        #
        # docId    - Identifier of the document to save.
        # fileName - Name of the Word file.
        # fmt      - Value of enumeration type WdSaveFormat (see wordConst.tcl).
        #            If not given or the empty string, the file is stored in the native
        #            format corresponding to the used Word version.
        #
        # No return value.
        #
        # See also: SaveAsPdf

        variable wordVersion

        set fileName [file nativename $fileName]
        set appId [::Cawt::GetApplicationId $docId]
        if { $fmt eq "" } {
            if { $wordVersion eq "14.0" } {
                $docId SaveAs $fileName [expr $::Word::wdFormatDocumentDefault]
            } else {
                $docId SaveAs $fileName
            }
        } else {
            ::Cawt::ShowAlerts $appId false
            $docId SaveAs $fileName $fmt
            ::Cawt::ShowAlerts $appId true
        }
    }

    proc SaveAsPdf { docId fileName } {
        # Save a document to a PDF file.
        #
        # docId    - Identifier of the document to export.
        # fileName - Name of the PDF file.
        #
        # PDF export is supported since Word 2007.
        # If your Word version is older an error is thrown.
        #
        # Note, that for Word 2007 you need the Microsoft Office Add-in
        # "Microsoft Save as PDF or XPS" available from
        # http://www.microsoft.com/en-us/download/details.aspx?id=7
        #
        # No return value.
        #
        # See also: SaveAs

        variable wordVersion

        if { $wordVersion < 12.0 } {
            error "PDF export available only in Word 2007 and up."
        }

        set fileName [file nativename $fileName]
        set appId [::Cawt::GetApplicationId $docId]

        ::Cawt::ShowAlerts $appId false
        $docId -callnamedargs ExportAsFixedFormat \
               OutputFileName $fileName \
               ExportFormat $::Word::wdExportFormatPDF \
               OpenAfterExport [::Cawt::TclBool false] \
               OptimizeFor $::Word::wdExportOptimizeForPrint \
               Range $::Word::wdExportAllDocument \
               From [expr 1] \
               To [expr 1] \
               Item $::Word::wdExportDocumentContent \
               IncludeDocProps [::Cawt::TclBool true] \
               KeepIRM [::Cawt::TclBool true] \
               CreateBookmarks $::Word::wdExportCreateHeadingBookmarks \
               DocStructureTags [::Cawt::TclBool true] \
               BitmapMissingFonts [::Cawt::TclBool true] \
               UseISO19005_1 [::Cawt::TclBool false]
        ::Cawt::ShowAlerts $appId true
    }

    proc SetCompatibilityMode { docId { mode $::Word::wdWord2010 } } {
        # Set the compatibility mode of a document.
        #
        # docId - Identifier of the document.
        # mode  - Compatibility mode of the document.
        #         Value of enumeration type WdCompatibilityMode (see wordConst.tcl).
        #
        # Available only for Word 2010 and up.
        #
        # No return value.

        variable wordVersion

        if { $wordVersion eq "14.0" } {
            $docId SetCompatibilityMode [expr $mode]
        }
    }

    proc AddDocument { appId { type "" } { visible true } } {
        # Add a new empty document to a Word instance.
        #
        # appId   - Identifier of the Word instance.
        # type    - Value of enumeration type WdNewDocumentType (see wordConst.tcl).
        # visible - true: Show the application window.
        #           false: Hide the application window.
        #
        # Return the identifier of the new document.
        #
        # See also: OpenDocument

        if { $type eq "" } {
            set type $::Word::wdNewBlankDocument
        }
        set docs [$appId Documents]
        # Add([Template], [NewTemplate], [DocumentType], [Visible]) As Document
        set docId [$docs -callnamedargs Add \
                         DocumentType $type \
                         Visible [::Cawt::TclInt $visible]]
        ::Cawt::Destroy $docs
        return $docId
    }

    proc GetNumDocuments { appId } {
        # Return the number of documents in a Word application.
        #
        # appId - Identifier of the Word instance.
        #
        # See also: AddDocument OpenDocument

        return [$appId -with { Documents } Count]
    }

    proc OpenDocument { appId fileName { readOnly false } } {
        # Open a document, i.e load a Word file.
        #
        # appId    - Identifier of the Word instance.
        # fileName - Name of the Word file.
        # readOnly - true: Open the document in read-only mode.
        #            false: Open the document in read-write mode.
        #
        # Return the identifier of the opened document. If the document was already open,
        # activate that document and return the identifier to that document.
        #
        # See also: AddDocument

        set nativeName  [file nativename $fileName]
        set docs [$appId Documents]
        set retVal [catch {[$docs Item [file tail $fileName]] Activate} d]
        if { $retVal == 0 } {
            puts "$nativeName already open"
            set docId [$docs Item [file tail $fileName]]
        } else {
            # Open(FileName, [ConfirmConversions], [ReadOnly],
            # [AddToRecentFiles], [PasswordDocument], [PasswordTemplate],
            # [Revert], [WritePasswordDocument], [WritePasswordTemplate],
            # [Format], [Encoding], [Visible], [OpenAndRepair],
            # [DocumentDirection], [NoEncodingDialog], [XMLTransform])
            # As Document
            set docId [$docs -callnamedargs Open \
                             FileName $nativeName \
                             ReadOnly [::Cawt::TclInt $readOnly]]
        }
        ::Cawt::Destroy $docs
        return $docId
    }

    proc GetDocumentIdByIndex { appId index } {
        # Find a document by it's index.
        #
        # appId - Identifier of the Word instance.
        # index - Index of the document to find.
        #
        # Return the identifier of the found document.
        # If the index is out of bounds an error is thrown.
        #
        # See also: GetNumDocuments GetDocumentName

        set count [::Word::GetNumDocuments $appId]

        if { $index < 1 || $index > $count } {
            error "GetDocumentIdByIndex: Invalid index $index given."
        }
        return [$appId -with { Documents } Item $index]
    }

    proc GetDocumentName { docId } {
        # Get the name of a document.
        #
        # docId - Identifier of the document.
        #
        # Return the name of the document (i.e. the full path name of the
        # corresponding Word file) as a string.

        return [$docId FullName]
    }

    proc AppendParagraph { docId { text "" } { spaceAfter -1 } } {
        # Append a paragraph at the end of the document.
        #
        # docId - Identifier of the document.
        # text  - Appended text string.
        #
        # Append a new paragraph to the end of the document and (optionally)
        # insert a text string in the new paragraph.
        #
        # Return the new end range.
        #
        # See also: AddParagraph

        set endRange [GetEndRange $docId]
        $endRange InsertParagraphAfter
        if { $text ne "" } {
            $endRange InsertAfter $text
        }
        if { $spaceAfter >= 0 } {
            $endRange -with { ParagraphFormat } SpaceAfter $spaceAfter
        }
        ::Cawt::Destroy $endRange
        return [GetEndRange $docId]
    }

    proc AddParagraph { rangeId { where "after" } } {
        # Add a new paragraph to a document.
        #
        # rangeId - Identifier of the text range.
        # where   - Insertion point of the new paragraph.
        #
        # No return value.
        #
        # See also: AppendParagraph

        if { $where eq "after" } {
            $rangeId InsertParagraphAfter
        } else {
            $rangeId InsertParagraphBefore
        }
    }

    proc AppendText { docId text } {
        # Append text to a Word document.
        #
        # docId - Identifier of the document.
        # text  - Appended text string.
        #
        # The text string is appended at the current end range of the document.
        #
        # See also:

        set endRange [::Word::GetEndRange $docId]
        set para [$docId -with { Content Paragraphs } Add $endRange]
        set range [$para Range]
        $range InsertAfter $text
        ::Cawt::Destroy $para
        ::Cawt::Destroy $endRange
        return $range
    }

    proc SetHyperlink { docId rangeId link { textDisplay "" } } {
        # Insert a hyperlink into a Word document.
        #
        # docId       - Identifier of the document.
        # rangeId     - Identifier of the text range.
        # link        - URL of the hyperlink.
        # textDisplay - Text to be displayed instead of the URL.
        #
        # URL's are specified as strings. "file://myLinkedFile" specifies a link
        # to a local file.
        #
        # No return value.

        if { $textDisplay eq "" } {
            set textDisplay $link
        }

        set hyperId [$docId Hyperlinks]
        # Add(Anchor As Object, [Address], [SubAddress], [ScreenTip],
        # [TextToDisplay], [Target]) As Hyperlink
        $hyperId -callnamedargs Add \
                 Anchor  $rangeId \
                 Address $link \
                 TextToDisplay $textDisplay
        ::Cawt::Destroy $hyperId
    }

   proc InsertImage { rangeId imgFileName { linkToFile false } { saveWithDocument false } } {
        # Insert an image into a range of a document.
        #
        # rangeId          - Identifier of the text range.
        # imgFileName      - File name of the image (as absolute path).
        # linkToFile       - Insert a link to the image instead of a copy.
        # saveWithDocument - If using a link, store the image in the file, too.
        #
        # The file name of the image must be an absolute pathname. Use a
        # construct like [file join [pwd] "myImage.gif"] to insert
        # images from the current directory.
        #
        # Return the identifier of the inserted image.
        #
        # See also: CropImage

	set fileName [file nativename $imgFileName]
        if { 0 } {
            # TODO
            set imgId [$rangeId -with { InlineShapes } AddPicture $fileName \
                       [::Cawt::TclInt $linkToFile] [::Cawt::TclInt $saveWithDocument]]
        } else {
            set imgId [$rangeId -with { InlineShapes } AddPicture $fileName]
        }
        return $imgId
    }

   proc CropImage { imgId { cropBottom 0.0 } { cropTop 0.0 } { cropLeft 0.0 } { cropRight 0.0 } } {
        # Crop an image at the four borders.
        #
        # imgId      - Identifier of the image.
        # cropBottom - Crop amount at the bottom border.
        # cropTop    - Crop amount at the top border.
        # cropLeft   - Crop amount at the left border.
        # cropRight  - Crop amount at the right border.
        #
        # The crop values must be specified in points.
        # Use ::Cawt::CentiMetersToPoints or ::Cawt::InchesToPoints
        # for conversion.
        #
        # No return value.
        #
        # See also: InsertImage ::Cawt::CentiMetersToPoints ::Cawt::InchesToPoints

        $imgId -with { PictureFormat } CropBottom $cropBottom
        $imgId -with { PictureFormat } CropTop    $cropTop
        $imgId -with { PictureFormat } CropLeft   $cropLeft
        $imgId -with { PictureFormat } CropRight  $cropRight
    }

    proc AddTable { docId rangeId numRows numCols { spaceAfter -1 } } {
        # Add a new table in a text range.
        #
        # docId      - Identifier of the document.
        # rangeId    - Identifier of the text range.
        # numRows    - Number of rows of the new table.
        # numCols    - Number of columns of the new table.
        # spaceAfter - Spacing in points after the table.
        #
        # Return the identifier of the new table.
        #
        # See also: GetNumRows GetNumColumns

        set tableId [$docId -with { Tables } Add $rangeId $numRows $numCols]
        if { $spaceAfter >= 0 } {
            $tableId -with { Range ParagraphFormat } SpaceAfter $spaceAfter
        }
        return $tableId
    }

    proc GetNumTables { docId } {
        # Return the number of tables of a Word document.
        #
        # docId - Identifier of the document.
        #
        # See also: AddTable GetNumRows GetNumCols

        return [$docId -with { Tables } Count]
    }

    proc GetTableIdByIndex { docId index } {
        # Find a table by it's index.
        #
        # docId - Identifier of the document.
        # index - Index of the table to find.
        #
        # Return the identifier of the found table.
        # If the index is out of bounds an error is thrown.
        #
        # See also: GetNumTables

        set count [::Word::GetNumTables $docId]

        if { $index < 1 || $index > $count } {
            error "GetTableIdByIndex: Invalid index $index given."
        }
        return [$docId -with { Tables } Item $index]
    }

    proc SetTableBorderLineStyle { tableId \
              { outsideLineStyle -1 } \
              { insideLineStyle  -1 } } {
        # Set the border line styles of a Word table.
        #
        # tableId          - Identifier of the Word table.
        # outsideLineStyle - Outside border style.
        # insideLineStyle  - Inside border style.
        #
        # The values of "outsideLineStyle" and "insideLineStyle" must
        # be of enumeration type WdLineStyle (see WordConst.tcl).
        #
        # See also: AddTable SetTableBorderLineWidth

        if { $outsideLineStyle < 0 } {
            set outsideLineStyle $::Word::wdLineStyleSingle
        }
        if { $insideLineStyle < 0 } {
            set insideLineStyle $::Word::wdLineStyleSingle
        }
        set border [$tableId Borders]
        $border OutsideLineStyle $outsideLineStyle
        $border InsideLineStyle  $insideLineStyle
        ::Cawt::Destroy $border
    }

    proc SetTableBorderLineWidth { tableId \
              { outsideLineWidth -1 } \
              { insideLineWidth  -1 } } {
        # Set the border line widths of a Word table.
        #
        # tableId          - Identifier of the Word table.
        # outsideLineWidth - Outside border line width.
        # insideLineWidth  - Inside border line width.
        #
        # The values of "outsideLineWidth" and "insideLineWidth" must
        # be of enumeration type WdLineWidth (see WordConst.tcl).
        #
        # See also: AddTable SetTableBorderLineStyle

        if { $outsideLineWidth < 0 } {
            set outsideLineWidth $::Word::wdLineWidth050pt
        }
        if { $insideLineWidth < 0 } {
            set insideLineWidth $::Word::wdLineWidth050pt
        }
        set border [$tableId Borders]
        $border OutsideLineWidth $outsideLineWidth
        $border InsideLineWidth  $insideLineWidth
        ::Cawt::Destroy $border
    }

    proc GetNumRows { tableId } {
        # Return the number of rows of a Word table.
        #
        # tableId - Identifier of the Word table.
        #
        # See also: GetNumColumns

        return [$tableId -with { Rows } Count]
    }

    proc GetNumColumns { tableId } {
        # Return the number of columns of a Word table.
        #
        # tableId - Identifier of the Word table.
        #
        # See also: GetNumRows

        return [$tableId -with { Columns } Count]
    }

    proc AddRow { tableId { beforeRowNum end } } {
        # Add a new row to a table.
        #
        # tableId      - Identifier of the Word table.
        # beforeRowNum - Insertion row number. Row numbering starts with 1.
        #                The new row is inserted before the given row number.
        #                If not specified or "end", the new row is appended at
        #                the end.
        #
        # No return value.
        #
        # See also: GetNumRows

        set rowsId [$tableId Rows]
        if { $beforeRowNum eq "end" } {
            $rowsId Add
        } else {
            if { $beforeRowNum < 1 || $beforeRowNum > [::Word::GetNumRows $tableId] } {
                error "AddRow: Invalid row number $beforeRowNum given."
            }
            set rowId [$tableId -with { Rows } Item $beforeRowNum]
            $rowsId Add $rowId
        }
    }

    proc GetCellRange { tableId row col } {
        # Return a cell of a Word table as a range.
        #
        # tableId - Identifier of the Word table.
        # row     - Row number. Row numbering starts with 1.
        # col     - Column number. Column numbering starts with 1.
        #
        # Return a range consisting of 1 cell of a Word table.
        #
        # See also: GetRowRange GetColumnRange

        set cellId  [$tableId Cell $row $col]
        set rangeId [$cellId Range]
        ::Cawt::Destroy $cellId
        return $rangeId
    }

    proc GetRowRange { tableId row } {
        # Return a row of a Word table as a range.
        #
        # tableId - Identifier of the Word table.
        # row     - Row number. Row numbering starts with 1.
        #
        # Return a range consisting of all cells of a row.
        #
        # See also: GetCellRange GetColumnRange

        set rowId [$tableId -with { Rows } Item $row]
        set rangeId [$rowId Range]
        ::Cawt::Destroy $rowId
        return $rangeId
    }

    proc GetColumnRange { tableId col } {
        # Return a column of a Word table as a selection.
        #
        # tableId - Identifier of the Word table.
        # col     - Column number. Column numbering starts with 1.
        #
        # Return a selection consisting of all cells of a column.
        # Note, that a selection is returned and not a range,
        # because columns do not have a range property.
        #
        # See also: GetCellRange GetRowRange

        set colId [$tableId -with { Columns } Item $col]
        $colId Select
        set selectId [$tableId -with { Application } Selection]
        $selectId SelectColumn
        ::Cawt::Destroy $colId
        return $selectId
    }

    proc SetCellValue { tableId row col val } {
        # Set the value of a Word table cell.
        #
        # tableId - Identifier of the Word table.
        # row     - Row number. Row numbering starts with 1.
        # col     - Column number. Column numbering starts with 1.
        # val     - String value of the cell.
        #
        # See also: GetCellValue SetRowValues SetMatrixValues

        set rangeId [::Word::GetCellRange $tableId $row $col]
        $rangeId Text $val
        ::Cawt::Destroy $rangeId
        return $rangeId
    }

    proc GetCellValue { tableId row col } {
        # Return the value of a Word table cell.
        #
        # tableId - Identifier of the Word table.
        # row     - Row number. Row numbering starts with 1.
        # col     - Column number. Column numbering starts with 1.
        #
        # Return the value of the specified cell as a string.
        #
        # See also: SetCellValue

        set rangeId [::Word::GetCellRange $tableId $row $col]
        set val [::Word::TrimString [$rangeId Text]]
        ::Cawt::Destroy $rangeId
        return $val
    }

    proc SetRowValues { tableId row valList { startCol 1 } { numVals 0 } } {
        # Insert row values from a Tcl list.
        #
        # tableId  - Identifier of the Word table.
        # row      - Row number. Row numbering starts with 1.
        # valList  - List of values to be inserted.
        # startCol - Column number of insertion start. Column numbering starts with 1.
        # numVals  - Negative or zero: All list values are inserted.
        #            Positive: numVals columns are filled with the list values
        #            (starting at list index 0).
        #
        # No return value. If valList is an empty list, an error is thrown.
        #
        # See also: GetRowValues SetColumnValues SetCellValue

        set len [llength $valList]
        if { $numVals > 0 } {
            if { $numVals < $len } {
                set len $numVals
            }
        }
        set ind 0
        for { set c $startCol } { $c < [expr {$startCol + $len}] } { incr c } {
            SetCellValue $tableId $row $c [lindex $valList $ind]
            incr ind
        }
    }

    proc GetRowValues { tableId row { startCol 1 } { numVals 0 } } {
        # Return row values of a Word table as a Tcl list.
        #
        # tableId  - Identifier of the Word table.
        # row      - Row number. Row numbering starts with 1.
        # startCol - Column number of start. Column numbering starts with 1.
        # numVals  - Negative or zero: All available row values are returned.
        #            Positive: Only numVals values of the row are returned.
        #
        # Return the values of the specified row or row range as a Tcl list.
        #
        # See also: SetRowValues GetColumnValues GetCellValue

        if { $numVals <= 0 } {
            set len [::Word::GetNumColumns $tableId]
        } else {
            set len $numVals
        }
        set valList [list]
        set col $startCol
        set ind 0
        while { $ind < $len } {
            set val [::Word::GetCellValue $tableId $row $col]
            lappend valList $val
            incr ind
            incr col
        }
        return $valList
    }

    proc SetColumnWidth { tableId col width } {
        # Set the width of a table column.
        #
        # tableId - Identifier of the Word table.
        # col     - Column number. Column numbering starts with 1.
        # width   - Column width of table column in points.
        #
        # No return value.
        #
        # See also: SetColumnsWidth InchesToPoints

        set colId [$tableId -with { Columns } Item $col]
        $colId Width $width
        ::Cawt::Destroy $colId
    }

    proc SetColumnsWidth { tableId startCol endCol width } {
        # Set the width of a range of table columns.
        #
        # tableId  - Identifier of the Word table.
        # startCol - Range start column number. Column numbering starts with 1.
        # endCol   - Range end column number. Column numbering starts with 1.
        # width    - Column width of table column in points.
        #
        # No return value.
        #
        # See also: SetColumnWidth InchesToPoints

        for { set c $startCol } { $c <= $endCol } { incr c } {
            SetColumnWidth $tableId $c $width
        }
    }

    proc SetColumnValues { tableId col valList { startRow 1 } { numVals 0 } } {
        # Insert column values into a Word table.
        #
        # tableId  - Identifier of the Word table.
        # col      - Column number. Column numbering starts with 1.
        # valList  - List of values to be inserted.
        # startRow - Row number of insertion start. Row numbering starts with 1.
        # numVals  - Negative or zero: All list values are inserted.
        #            Positive: numVals rows are filled with the list values
        #            (starting at list index 0).
        #
        # No return value.
        #
        # See also: GetColumnValues SetRowValues SetCellValue

        set len [llength $valList]
        if { $numVals > 0 } {
            if { $numVals < $len } {
                set len $numVals
            }
        }
        set ind 0
        for { set r $startRow } { $r < [expr {$startRow + $len}] } { incr r } {
            SetCellValue $tableId $r $col [lindex $valList $ind]
            incr ind
        }
    }

    proc GetColumnValues { tableId col { startRow 1 } { numVals 0 } } {
        # Return column values of a Word table as a Tcl list.
        #
        # tableId  - Identifier of the Word table.
        # col      - Column number. Column numbering starts with 1.
        # startRow - Row number of start. Row numbering starts with 1.
        # numVals  - Negative or zero: All available column values are returned.
        #            Positive: Only numVals values of the column are returned.
        #
        # Return the values of the specified column or column range as a Tcl list.
        #
        # See also: SetColumnValues GetRowValues GetCellValue

        if { $numVals <= 0 } {
            set len [GetNumRows $tableId]
        } else {
            set len $numVals
        }
        set valList [list]
        set row $startRow
        set ind 0
        while { $ind < $len } {
            set val [GetCellValue $tableId $row $col]
            if { $val eq "" } {
                set val2 [GetCellValue $tableId [expr {$row+1}] $col]
                if { $val2 eq "" } {
                    break
                }
            }
            lappend valList $val
            incr ind
            incr row
        }
        return $valList
    }
}
