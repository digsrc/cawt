# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval Word {

    namespace ensemble create

    namespace export AddBookmark
    namespace export AddDocument
    namespace export AddPageBreak
    namespace export AddParagraph
    namespace export AddRow
    namespace export AddTable
    namespace export AddText
    namespace export AppendParagraph
    namespace export AppendText
    namespace export Close
    namespace export ConfigureCaption
    namespace export CreateRange
    namespace export CreateRangeAfter
    namespace export CropImage
    namespace export ExtendRange
    namespace export FindString
    namespace export GetBookmarkName
    namespace export GetCellRange
    namespace export GetCellValue
    namespace export GetColumnRange
    namespace export GetColumnValues
    namespace export GetCompatibilityMode
    namespace export GetDocumentId
    namespace export GetDocumentIdByIndex
    namespace export GetDocumentName
    namespace export GetEndRange
    namespace export GetExtString
    namespace export GetListGalleryId
    namespace export GetListTemplateId
    namespace export GetNumCharacters
    namespace export GetNumColumns
    namespace export GetNumDocuments
    namespace export GetNumRows
    namespace export GetNumTables
    namespace export GetRangeEndIndex
    namespace export GetRangeInformation
    namespace export GetRangeStartIndex
    namespace export GetRowRange
    namespace export GetRowValues
    namespace export GetSelectionRange
    namespace export GetStartRange
    namespace export GetTableIdByIndex
    namespace export GetVersion
    namespace export InsertCaption
    namespace export InsertFile
    namespace export InsertImage
    namespace export InsertList
    namespace export InsertText
    namespace export Open
    namespace export OpenDocument
    namespace export OpenNew
    namespace export PrintRange
    namespace export Quit
    namespace export ReplaceByProc
    namespace export ReplaceString
    namespace export SaveAs
    namespace export SaveAsPdf
    namespace export ScaleImage
    namespace export SelectRange
    namespace export SetCellValue
    namespace export SetColumnValues
    namespace export SetColumnWidth
    namespace export SetColumnsWidth
    namespace export SetCompatibilityMode
    namespace export SetHyperlink
    namespace export SetInternalHyperlink
    namespace export SetLinkToBookmark
    namespace export SetRangeBackgroundColor
    namespace export SetRangeBackgroundColorByEnum
    namespace export SetRangeEndIndex
    namespace export SetRangeFontBold
    namespace export SetRangeFontItalic
    namespace export SetRangeFontName
    namespace export SetRangeFontSize
    namespace export SetRangeFontUnderline
    namespace export SetRangeHighlightColorByEnum
    namespace export SetRangeHorizontalAlignment
    namespace export SetRangeStartIndex
    namespace export SetRangeStyle
    namespace export SetRowValues
    namespace export SetTableBorderLineStyle
    namespace export SetTableBorderLineWidth
    namespace export ToggleSpellCheck
    namespace export TrimString
    namespace export UpdateFields
    namespace export Visible

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

    proc _IsDocument { objId } {
        # ActiveTheme is a property of the Word Document class.
        set retVal [catch {$objId ActiveTheme} errMsg]
        if { $retVal == 0 } {
            return true
        } else {
            return false
        }
    }

    proc _FindOrReplace { objId mode searchStr matchCase { replaceStr "" } { howMuch "one" } } {
        set myFind [$objId Find]

        # Execute([FindText], [MatchCase], [MatchWholeWord], [MatchWildcards],
        # [MatchSoundsLike], [MatchAllWordForms], [Forward], [Wrap], [Format],
        # [ReplaceWith], [Replace], [MatchKashida], [MatchDiacritics],
        # [MatchAlefHamza], [MatchControl]) As Boolean
        if { $mode eq "find" } {
            set retVal [$myFind -callnamedargs Execute \
                                FindText $searchStr \
                                MatchCase [Cawt TclBool $matchCase] \
                                Wrap $Word::wdFindStop \
                                Forward True]
        } else {
            set howMuchEnum $Word::wdReplaceOne
            if { $howMuch ne "one" } {
                set howMuchEnum $Word::wdReplaceAll
            }
            set retVal [$myFind -callnamedargs Execute \
                                FindText $searchStr \
                                ReplaceWith $replaceStr \
                                Replace $howMuchEnum \
                                Wrap $Word::wdFindStop \
                                MatchCase [Cawt TclBool $matchCase] \
                                Forward True]
        }
        Cawt Destroy $myFind
        return $retVal
    }

    proc FindString { rangeOrDocId str { matchCase true } } {
        # Find a string in a text range or a document.
        #
        # rangeOrDocId - Identifier of a text range or a document identifier.
        # str          - Search string.
        # matchCase    - Flag indicating case sensitive search.
        #
        # Return true, if string was found. Otherwise false.
        # If the string was found, the selection is set to the found string.
        #
        # See also: ReplaceString ReplaceByProc GetSelectionRange

        if { [Word::_IsDocument $rangeOrDocId] } {
            set numFound 0
            set stories [$rangeOrDocId StoryRanges]
            $stories -iterate story {
                lappend storyList $story
                set retVal [Word::_FindOrReplace $story "find" $str $matchCase]
                incr numFound
                set nextStory [$story NextStoryRange]
                while { [Cawt IsComObject $nextStory] } {
                    lappend storyList $nextStory
                    set retVal [Word::_FindOrReplace $nextStory "find" $str $matchCase]
                    incr numFound
                    set nextStory [$nextStory NextStoryRange]
                }
            }
            foreach story $storyList {
                Cawt Destroy $story
            }
            Cawt Destroy $stories
            return $numFound
        } else {
            return [Word::_FindOrReplace $rangeOrDocId "find" $str $matchCase]
        }
    }

    proc ReplaceString { rangeOrDocId searchStr replaceStr \
                        { howMuch "one" } { matchCase true } } {
        # Replace a string in a text range or a document. Simple case.
        #
        # rangeOrDocId - Identifier of a text range or a document identifier.
        # searchStr    - Search string.
        # replaceStr   - Replacement string.
        # howMuch      - "one" to replace first occurence only. "all" to replace all occurences.
        # matchCase    - Flag indicating case sensitive search.
        #
        # Return true, if string could be replaced, i.e. the search string was found.
        # Otherwise false.
        #
        # See also: FindString ReplaceByProc

        if { [Word::_IsDocument $rangeOrDocId] } {
            set numReplaced 0
            set stories [$rangeOrDocId StoryRanges]
            $stories -iterate story {
                lappend storyList $story
                set retVal [Word::_FindOrReplace $story "replace" $searchStr $matchCase $replaceStr $howMuch]
                incr numReplaced
                set nextStory [$story NextStoryRange]
                while { [Cawt IsComObject $nextStory] } {
                    lappend storyList $nextStory
                    set retVal [Word::_FindOrReplace $nextStory "replace" $searchStr $matchCase $replaceStr $howMuch]
                    incr numReplaced
                    set nextStory [$nextStory NextStoryRange]
                }
            }
            foreach story $storyList {
                Cawt Destroy $story
            }
            Cawt Destroy $stories
            return $numReplaced
        } else {
            return [Word::_FindOrReplace $rangeOrDocId "replace" $searchStr $matchCase $replaceStr $howMuch]
        }
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
        # See also: FindString ReplaceString

        set myFind [$rangeId Find]
        set count 0
        while { 1 } {
            # See proc _FindOrReplace for a parameter list of the Execute command.
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
        Cawt Destroy $myFind
    }

    proc GetNumCharacters { docId } {
        # Return the number of characters in a Word document.
        #
        # docId - Identifier of the document.
        #
        # See also: GetNumDocuments GetNumTables GetNumCharacters

        return [$docId -with { Characters } Count]
    }

    proc CreateRange { docId startIndex endIndex } {
        # Create a new text range.
        #
        # docId      - Identifier of the document.
        # startIndex - The start index of the range in characters.
        # endIndex   - The end index of the range in characters.
        #
        # Return the identifier of the new text range.
        #
        # See also: CreateRangeAfter SelectRange GetSelectionRange

        return [$docId Range $startIndex $endIndex]
    }

    proc CreateRangeAfter { rangeId } {
        # Create a new text range after specified range.
        #
        # rangeId - Identifier of the text range.
        #
        # Return the identifier of the new text range.
        #
        # See also: CreateRange SelectRange GetSelectionRange

        set docId [Word GetDocumentId $rangeId]
        set index [Word GetRangeEndIndex $rangeId]
        set rangeId [Word CreateRange $docId $index $index]
        Cawt Destroy $docId
        return $rangeId
    }

    proc SelectRange { rangeId } {
        # Select a text range.
        #
        # rangeId - Identifier of the text range.
        #
        # No return value.
        #
        # See also: GetSelectionRange

        $rangeId Select
    }

    proc GetSelectionRange { docId } {
        # Return the text range representing the current selection.
        #
        # docId - Identifier of the document.
        #
        # See also: GetStartRange GetEndRange SelectRange

        return [$docId -with { ActiveWindow } Selection]
    }

    proc GetStartRange { docId } {
        # Return a text range representing the start of the document.
        #
        # docId - Identifier of the document.
        #
        # See also: CreateRange GetSelectionRange GetEndRange

        return [Word CreateRange $docId 0 0]
    }

    proc GetEndRange { docId } {
        # Return the text range representing the end of the document.
        #
        # docId - Identifier of the document.
        #
        # Note: This corresponds to the built-in bookmark \endofdoc.
        #       The end range of an empty document is (0, 0), although
        #       GetNumCharacters returns 1.
        #
        # See also: GetSelectionRange GetStartRange GetNumCharacters

        set bookMarks [$docId Bookmarks]
        set endOfDoc  [$bookMarks Item "\\endofdoc"]
        set endRange  [$endOfDoc Range]
        Cawt Destroy $endOfDoc
        Cawt Destroy $bookMarks
        set endIndex [Word GetRangeEndIndex $endRange]
        Cawt Destroy $endRange
        return [Word CreateRange $docId $endIndex $endIndex]
    }

    proc GetRangeInformation { rangeId type } {
        # Get information about a text range.
        #
        # rangeId - Identifier of the text range.
        # type    - Value of enumeration type WdInformation (see wordConst.tcl).
        #
        # Return the range information associated with the supplied type.
        #
        # See also: GetStartRange GetEndRange PrintRange

        return [$rangeId Information [Word GetEnum $type]]
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
              [Word GetRangeStartIndex $rangeId] [Word GetRangeEndIndex $rangeId]]
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

    proc SetRangeStartIndex { rangeId index } {
        # Set the start index of a text range.
        #
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

    proc SetRangeEndIndex { rangeId index } {
        # Set the end index of a text range.
        #
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
            set docId [Word GetDocumentId $rangeId]
            set index [$docId End]
            Cawt Destroy $docId
        }
        $rangeId End $index
    }

    proc ExtendRange { rangeId { startIncr 0 } { endIncr 0 } } {
        # Extend the range indices of a text range.
        #
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

        set startIndex [Word GetRangeStartIndex $rangeId]
        set endIndex   [Word GetRangeEndIndex   $rangeId]
        if { [string is integer $startIncr] } {
            set startIndex [expr $startIndex + $startIncr]
        } elseif { $startIncr eq "begin" } {
            set startIndex 0
        }
        if { [string is integer $endIncr] } {
            set endIndex [expr $endIndex + $endIncr]
        } elseif { $endIncr eq "end" } {
            set docId [Word GetDocumentId $rangeId]
            set endRange [GetEndRange $docId]
            set endIndex [$endRange End]
            Cawt Destroy $endRange
            Cawt Destroy $docId
        }
        $rangeId Start $startIndex
        $rangeId End $endIndex
        return $rangeId
    }

    proc SetRangeStyle { rangeId style } {
        # Set the style of a text range.
        #
        # rangeId - Identifier of the text range.
        # style   - Value of enumeration type WdBuiltinStyle (see wordConst.tcl).
        #           Often used values: Word::wdStyleHeading1, Word::wdStyleNormal
        #
        # No return value.
        #
        # See also: SetRangeFontSize SetRangeFontName

        set docId [Word GetDocumentId $rangeId]
        set styleId [$docId -with { Styles } Item [Word GetEnum $style]]
        $rangeId Style $styleId
        Cawt Destroy $styleId
        Cawt Destroy $docId
    }

    proc SetRangeFontName { rangeId fontName } {
        # Set the font name of a text range.
        #
        # rangeId  - Identifier of the text range.
        # fontName - Font name.
        #
        # No return value.
        #
        # See also: SetRangeFontSize SetRangeFontBold SetRangeFontItalic SetRangeFontUnderline

        $rangeId -with { Font } Name $fontName
    }

    proc SetRangeFontSize { rangeId fontSize } {
        # Set the font size of a text range.
        #
        # rangeId  - Identifier of the text range.
        # fontSize - Font size in points.
        #
        # No return value.
        #
        # See also: SetRangeFontName SetRangeFontBold SetRangeFontItalic SetRangeFontUnderline

        $rangeId -with { Font } Size [expr int($fontSize)]
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
        # See also: SetRangeFontName SetRangeFontSize SetRangeFontItalic SetRangeFontUnderline

        $rangeId -with { Font } Bold [Cawt TclInt $onOff]
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
        # See also: SetRangeFontName SetRangeFontSize SetRangeFontBold SetRangeFontUnderline

        $rangeId -with { Font } Italic [Cawt TclInt $onOff]
    }

    proc SetRangeFontUnderline { rangeId { onOff true } { color wdColorAutomatic } } {
        # Toggle the underline font style of a text range.
        #
        # rangeId - Identifier of the text range.
        # onOff   - true: Set underline style on.
        #           false: Set underline style off.
        # color   - Value of enumeration type WdColor (see wordConst.tcl)
        #
        # No return value.
        #
        # See also: SetRangeFontName SetRangeFontSize SetRangeFontBold SetRangeFontItalic

        $rangeId -with { Font } Underline [Cawt TclInt $onOff]
        if { $onOff } {
            $rangeId -with { Font } UnderlineColor [Word GetEnum $color]
        }
    }

    proc SetRangeHorizontalAlignment { rangeId align } {
        # Set the horizontal alignment of a text range.
        #
        # rangeId - Identifier of the text range.
        # align   - Value of enumeration type WdParagraphAlignment (see wordConst.tcl)
        #           or any of the following strings: left, right, center.
        #
        # No return value.
        #
        # See also: SetRangeHighlightColorByEnum

        if { $align eq "center" } {
            set alignEnum $Word::wdAlignParagraphCenter
        } elseif { $align eq "left" } {
            set alignEnum $Word::wdAlignParagraphLeft
        } elseif { $align eq "right" } {
            set alignEnum $Word::wdAlignParagraphRight
        } else {
            set alignEnum [Word GetEnum $align]
        }

        $rangeId -with { ParagraphFormat } Alignment $alignEnum
    }

    proc SetRangeHighlightColorByEnum { rangeId colorEnum } {
        # Set the highlight color of a text range.
        #
        # rangeId   - Identifier of the text range.
        # colorEnum - Value of enumeration type WdColorIndex (see wordConst.tcl).
        #
        # No return value.
        #
        # See also: SetRangeBackgroundColorByEnum

        $rangeId HighlightColorIndex [Word GetEnum $colorEnum]
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

        $rangeId -with { Cells Shading } BackgroundPatternColor [Word GetEnum $colorEnum]
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
                                   [Cawt RgbToColor $r $g $b]
    }

    proc AddPageBreak { rangeId } {
        # Add a page break to a text range.
        #
        # rangeId - Identifier of the text range.
        #
        # No return value.
        #
        # See also: AddParagraph

        $rangeId Collapse $Word::wdCollapseEnd
        $rangeId InsertBreak [expr { int ($Word::wdPageBreak) }]
        $rangeId Collapse $Word::wdCollapseEnd
    }

    proc AddBookmark { rangeId name } {
        # Add a bookmark to a text range.
        #
        # rangeId - Identifier of the text range.
        # name    - Name of the bookmark. 
        #
        # Return the bookmark identifier.
        #
        # See also: SetLinkToBookmark GetBookmarkName

        set docId [Word GetDocumentId $rangeId]
        set bookmarks [$docId Bookmarks]
        # Create valid bookmark names.
        set validName [regsub -all { } $name {_}]
        set validName [regsub -all -- {-} $validName {_}]
        set bookmarkId [$bookmarks Add $validName $rangeId]

        Cawt Destroy $bookmarks
        Cawt Destroy $docId
        return $bookmarkId
    }

    proc GetBookmarkName { bookmarkId } {
        # Get the name of a bookmark.
        #
        # bookmarkId - Identifier of the boormark.
        #
        # Return the name of the bookmark.
        #
        # See also: AddBookmark SetLinkToBookmark

        return [$bookmarkId Name]
    }

    proc GetListGalleryId { appId galleryType } {
        # Get one of the 3 predefined list galleries.
        #
        # appId       - Identifier of the Word instance.
        # galleryType - Value of enumeration type WdListGalleryType (see wordConst.tcl).
        #
        # Return the identifier of the specified list gallery.
        #
        # See also: GetListTemplateId InsertList

        return [$appId -with { ListGalleries } Item [Word GetEnum $galleryType]]
    }

    proc GetListTemplateId { galleryId listType } {
        # Get one of the 7 predefined list templates.
        #
        # galleryId - Identifier of the Word gallery.
        # listType  - Value of enumeration type WdListType (see wordConst.tcl)
        #
        # Return the identifier of the specified list template.
        #
        # See also: GetListGalleryId InsertList

        return [$galleryId -with { ListTemplates } Item [Word GetEnum $listType]]
    }

    proc InsertList { rangeId stringList \
                      { galleryType wdBulletGallery } \
                      { listType wdListListNumOnly } } {
        # Insert a Word list.
        #
        # rangeId     - Identifier of the text range.
        # stringList  - List of text strings building up the Word list. 
        # galleryType - Value of enumeration type WdListGalleryType (see wordConst.tcl).
        # listType    - Value of enumeration type WdListType (see wordConst.tcl)
        #
        # Return the range of the Word list.
        #
        # See also: GetListGalleryId GetListTemplateId InsertCaption InsertFile InsertImage InsertText

        foreach line $stringList {
            append listStr "$line\n"
        }
        set appId [Cawt GetApplicationId $rangeId]
        set listRangeId [Word AddText $rangeId $listStr]
        set listGalleryId  [Word GetListGalleryId $appId $galleryType]
        set listTemplateId [Word GetListTemplateId $listGalleryId $listType]
        $listRangeId -with { ListFormat } ApplyListTemplate $listTemplateId
        Cawt Destroy $listTemplateId
        Cawt Destroy $listGalleryId
        Cawt Destroy $appId
        return $listRangeId
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
        #
        # See also: GetCompatibilityMode GetExtString

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
                return "Unknown Word version $version"
            }
        } else {
            return $version
        }
    }

    proc GetCompatibilityMode { appId { version "" } } {
        # Return the compatibility version of a Word application.
        #
        # appId   - Identifier of the Word instance.
        # version - Word version number.
        #
        # Return the compatibility mode of the current Word application, if 
        # version is not specified or the empty string.
        # If version is a valid Word version as returned by GetVersion, the
        # corresponding compatibility mode is returned.
        #
        # Note: The compatibility mode is a value of enumeration WdCompatibilityMode.
        #
        # See also: GetVersion GetExtString

        if { $version eq "" } {
            return $Word::wdCurrent
        } else {
            array set map {
                "11.0" $Word::wdWord2003
                "12.0" $Word::wdWord2007
                "14.0" $Word::wdWord2010
                "15.0" $Word::wdWord2013
            }
            if { [info exists map($version)] } {
                return $map($version)
            } else {
                error "Unknown Word version $version"
            }
        }
    }

    proc GetExtString { appId } {
        # Return the default extension of a Word file.
        #
        # appId - Identifier of the Word instance.
        #
        # Starting with Word 12 (2007) this is the string ".docx".
        # In previous versions it was ".doc".
        #
        # See also: GetCompatibilityMode GetVersion

        # appId is only needed, so we are sure, that wordVersion is initialized.
 
        variable wordVersion

        if { $wordVersion >= 12.0 } {
            return ".docx"
        } else {
            return ".doc"
        }
    }

    proc ToggleSpellCheck { appId onOff } {
        # Toggle checking of grammatical and spelling errors.
        #
        # appId - Identifier of the Word instance.
        #
        # No return value.
        #
        # See also: 

        $appId -with { ActiveDocument } ShowGrammaticalErrors [Cawt TclBool $onOff]
        $appId -with { ActiveDocument } ShowSpellingErrors    [Cawt TclBool $onOff]
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

        set appId [Cawt GetOrCreateApp $wordAppName false]
        set wordVersion [Word GetVersion $appId]
        Word Visible $appId $visible
        if { $width >= 0 } {
            $appId Width [expr $width]
        }
        if { $height >= 0 } {
            $appId Height [expr $height]
        }
        return $appId
    }

    proc Open { { visible true } { width -1 } { height -1 } } {
        # Open a Word instance. Use an already running instance, if available.
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

        set appId [Cawt GetOrCreateApp $wordAppName true]
        set wordVersion [Word GetVersion $appId]
        Word Visible $appId $visible
        if { $width >= 0 } {
            $appId Width [expr $width]
        }
        if { $height >= 0 } {
            $appId Height [expr $height]
        }
        return $appId
    }

    proc Quit { appId { showAlert true } } {
        # Quit a Word instance.
        #
        # appId     - Identifier of the Word instance.
        # showAlert - true: Show an alert window, if there are unsaved changes.
        #             false: Quit without saving any changes.
        #
        # No return value.
        #
        # See also: Open OpenNew

        if { ! $showAlert } {
            Cawt ShowAlerts $appId false
        }
        $appId Quit
    }

    proc Visible { appId visible } {
        # Toggle the visibility of a Word application window.
        #
        # appId   - Identifier of the Word instance.
        # visible - true: Show the application window.
        #           false: Hide the application window.
        #
        # No return value.
        #
        # See also: Open OpenNew

        $appId Visible [Cawt TclInt $visible]
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

        $docId Close [Cawt TclBool false]
    }

    proc UpdateFields { docId } {
        # Update all fields as well as tables of content and figures of a document.
        #
        # docId - Identifier of the document.
        #
        # No return value.
        #
        # See also: SaveAs

        set stories [$docId StoryRanges]
        $stories -iterate story {
            lappend storyList $story
            $story -with { Fields } Update
            set nextStory [$story NextStoryRange]
            while { [Cawt IsComObject $nextStory] } {
                lappend storyList $nextStory
                $nextStory -with { Fields } Update
                set nextStory [$nextStory NextStoryRange]
            }
        }
        foreach story $storyList {
            Cawt Destroy $story
        }
        Cawt Destroy $stories

        set tocs [$docId TablesOfContents]
        $tocs -iterate toc {
            $toc Update
            Cawt Destroy $toc
        }
        Cawt Destroy $tocs

        set tofs [$docId TablesOfFigures]
        $tofs -iterate tof {
            $tof Update
            Cawt Destroy $tof
        }
        Cawt Destroy $tofs
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
        set appId [Cawt GetApplicationId $docId]
        Cawt ShowAlerts $appId false
        if { $fmt eq "" } {
            if { $wordVersion >= 14.0 } {
                $docId SaveAs $fileName [expr $Word::wdFormatDocumentDefault]
            } else {
                $docId SaveAs $fileName
            }
        } else {
            $docId SaveAs $fileName [Word GetEnum $fmt]
        }
        Cawt ShowAlerts $appId true
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
        set appId [Cawt GetApplicationId $docId]

        Cawt ShowAlerts $appId false
        $docId -callnamedargs ExportAsFixedFormat \
               OutputFileName $fileName \
               ExportFormat $Word::wdExportFormatPDF \
               OpenAfterExport [Cawt TclBool false] \
               OptimizeFor $Word::wdExportOptimizeForPrint \
               Range $Word::wdExportAllDocument \
               From [expr 1] \
               To [expr 1] \
               Item $Word::wdExportDocumentContent \
               IncludeDocProps [Cawt TclBool true] \
               KeepIRM [Cawt TclBool true] \
               CreateBookmarks $Word::wdExportCreateHeadingBookmarks \
               DocStructureTags [Cawt TclBool true] \
               BitmapMissingFonts [Cawt TclBool true] \
               UseISO19005_1 [Cawt TclBool false]
        Cawt ShowAlerts $appId true
    }

    proc SetCompatibilityMode { docId { mode wdWord2010 } } {
        # Set the compatibility mode of a document.
        #
        # docId - Identifier of the document.
        # mode  - Compatibility mode of the document.
        #         Value of enumeration type WdCompatibilityMode (see wordConst.tcl).
        #
        # Available only for Word 2010 and up.
        #
        # No return value.
        #
        # See also: GetCompatibilityMode

        variable wordVersion

        if { $wordVersion >= 14.0 } {
            $docId SetCompatibilityMode [Word GetEnum $mode]
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
            set type $Word::wdNewBlankDocument
        }
        set docs [$appId Documents]
        # Add([Template], [NewTemplate], [DocumentType], [Visible]) As Document
        set docId [$docs -callnamedargs Add \
                         DocumentType [Word GetEnum $type] \
                         Visible [Cawt TclInt $visible]]
        Cawt Destroy $docs
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
        # Open a document, i.e. load a Word file.
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
                             ConfirmConversions [Cawt TclBool false] \
                             ReadOnly [Cawt TclInt $readOnly]]
        }
        Cawt Destroy $docs
        return $docId
    }

    proc GetDocumentIdByIndex { appId index } {
        # Find a document by its index.
        #
        # appId - Identifier of the Word instance.
        # index - Index of the document to find.
        #
        # Return the identifier of the found document.
        # If the index is out of bounds an error is thrown.
        #
        # See also: GetNumDocuments GetDocumentName

        set count [Word GetNumDocuments $appId]

        if { $index < 1 || $index > $count } {
            error "GetDocumentIdByIndex: Invalid index $index given."
        }
        return [$appId -with { Documents } Item $index]
    }

    proc GetDocumentId { componentId } {
        # Get the document identifier of a Word component.
        #
        # componentId - The identifier of a Word component.
        #
        # Word components having the Document property are ex. ranges, panes.

        return [$componentId Document]
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

    proc AppendParagraph { docId { spaceAfter -1 } } {
        # Append a paragraph at the end of the document.
        #
        # docId      - Identifier of the document.
        # spaceAfter - Spacing in points after the range.
        #
        # Append a new paragraph to the end of the document.
        #
        # No return value.
        #
        # See also: GetEndRange AddParagraph

        set endRange [Word GetEndRange $docId]
        $endRange InsertParagraphAfter
        if { $spaceAfter >= 0 } {
            $endRange -with { ParagraphFormat } SpaceAfter $spaceAfter
        }
        return $endRange
    }

    proc AddParagraph { rangeId { spaceAfter -1 } } {
        # Add a new paragraph to a document.
        #
        # rangeId    - Identifier of the text range.
        # spaceAfter - Spacing in points after the range.
        #
        # Return the new extended range.
        #
        # See also: AppendParagraph

        $rangeId InsertParagraphAfter
        if { $spaceAfter >= 0 } {
            $rangeId -with { ParagraphFormat } SpaceAfter $spaceAfter
        }
        return $rangeId
    }

    proc InsertText { docId text { addParagraph false } { style wdStyleNormal } } {
        # Insert text in a Word document.
        #
        # docId        - Identifier of the document.
        # text         - Text string to be inserted.
        # addParagraph - Add a paragraph after the text.
        # style        - Value of enumeration type WdBuiltinStyle (see wordConst.tcl).
        #
        # The text string is inserted at the start of the document with given style.
        #
        # Return the new text range.
        #
        # See also: AddText AppendText AddParagraph SetRangeStyle 
        #           InsertCaption InsertFile InsertImage InsertList

        set newRange [Word CreateRange $docId 0 0]
        $newRange InsertAfter $text
        if { $addParagraph } {
            $newRange InsertParagraphAfter
        }
        Word SetRangeStyle $newRange $style
        return $newRange
    }

    proc AppendText { docId text { addParagraph false } { style wdStyleNormal } } {
        # Append text to a Word document.
        #
        # docId        - Identifier of the document.
        # text         - Text string to be appended.
        # addParagraph - Add a paragraph after the text.
        # style        - Value of enumeration type WdBuiltinStyle (see wordConst.tcl).
        #
        # The text string is appended at the end of the document with given style.
        #
        # Return the new text range.
        #
        # See also: GetEndRange AddText InsertText AppendParagraph SetRangeStyle

        set newRange [Word GetEndRange $docId]
        $newRange InsertAfter $text
        if { $addParagraph } {
            $newRange InsertParagraphAfter
        }
        Word SetRangeStyle $newRange $style
        return $newRange
    }

    proc AddText { rangeId text { addParagraph false } { style wdStyleNormal } } {
        # Add text to a Word document.
        #
        # rangeId      - Identifier of the text range.
        # text         - Text string to be added.
        # addParagraph - Add a paragraph after the text.
        # style        - Value of enumeration type WdBuiltinStyle (see wordConst.tcl).
        #
        # The text string is appended to the supplied text range with given style.
        # Return the new text range.
        #
        # See also: AddText InsertText AppendParagraph SetRangeStyle

        set newStartIndex [$rangeId End]
        set docId [Word GetDocumentId $rangeId]
        set newRange [Word CreateRange $docId $newStartIndex $newStartIndex]
        $newRange InsertAfter $text
        if { $addParagraph } {
            $newRange InsertParagraphAfter
        }
        Word SetRangeStyle $newRange $style
        Cawt Destroy $docId
        return $newRange
    }

    proc SetHyperlink { rangeId link { textDisplay "" } } {
        # Insert an external hyperlink into a Word document.
        #
        # rangeId     - Identifier of the text range.
        # link        - URL of the hyperlink.
        # textDisplay - Text to be displayed instead of the URL.
        #
        # # URL's are specified as strings:
        # file://myLinkedFile specifies a link to a local file.
        # http://myLinkedWebpage specifies a link to a web address.
        #
        # No return value.
        #
        # See also: SetLinkToBookmark SetInternalHyperlink

        if { $textDisplay eq "" } {
            set textDisplay $link
        }

        set docId [Word GetDocumentId $rangeId]
        set hyperlinks [$docId Hyperlinks]
        # Add(Anchor As Object, [Address], [SubAddress], [ScreenTip],
        # [TextToDisplay], [Target]) As Hyperlink
        set hyperlink [$hyperlinks -callnamedargs Add \
                 Anchor  $rangeId \
                 Address $link \
                 TextToDisplay $textDisplay]
        Cawt Destroy $hyperlink
        Cawt Destroy $hyperlinks
        Cawt Destroy $docId
    }

    proc SetInternalHyperlink { rangeId subAddress { textDisplay "" } } {
        # Insert an internal hyperlink into a Word document.
        #
        # rangeId     - Identifier of the text range.
        # subAddress  - Internal reference.
        # textDisplay - Text to be displayed instead of the URL.
        #
        # No return value.
        #
        # See also: SetLinkToBookmark SetHyperlink

        if { $textDisplay eq "" } {
            set textDisplay $subAddress
        }

        set docId [Word GetDocumentId $rangeId]
        set hyperlinks [$docId Hyperlinks]
        # Add(Anchor As Object, [Address], [SubAddress], [ScreenTip],
        # [TextToDisplay], [Target]) As Hyperlink
        $hyperlinks -callnamedargs Add \
                 Anchor  $rangeId \
                 SubAddress $subAddress \
                 TextToDisplay $textDisplay
        Cawt Destroy $hyperlinks
        Cawt Destroy $docId
    }

    proc SetLinkToBookmark { rangeId bookmarkId { textDisplay "" } } {
        # Insert an internal link to a bookmark into a Word document.
        #
        # rangeId     - Identifier of the text range.
        # bookmarkId  - Identifier of the bookmark to link to.
        # textDisplay - Text to be displayed instead of the bookmark name.
        #
        # No return value.
        #
        # See also: AddBookmark GetBookmarkName SetHyperlink SetInternalHyperlink

        set bookmarkName [Word GetBookmarkName $bookmarkId]
        if { $textDisplay eq "" } {
            set textDisplay $bookmarkName
        }

        set docId [Word GetDocumentId $rangeId]
        set hyperlinks [$docId Hyperlinks]
        # Add(Anchor As Object, [Address], [SubAddress], [ScreenTip],
        # [TextToDisplay], [Target]) As Hyperlink
        $hyperlinks -callnamedargs Add \
                 Anchor        $rangeId \
                 Address       "" \
                 SubAddress    $bookmarkName \
                 TextToDisplay $textDisplay
        Cawt Destroy $hyperlinks
        Cawt Destroy $docId
    }

    proc InsertFile { rangeId fileName { pasteFormat "" } } {
        # Insert a file into a Word document.
        #
        # rangeId     - Identifier of the text range.
        # fileName    - Name of the file to insert.
        # pasteFormat - Value of enumeration type WdRecoveryType (see wordConst.tcl).
        #
        # Insert an external file a the text range identified by rangeId. If pasteFormat is
        # not specified or an empty string, the method InsertFile is used.
        # Otherwise the external file is opened in a new Word document, copied to the clipboard
        # and pasted into the text range. For pasting the PasteAndFormat method is used, so it is 
        # possible to merge the new text from the external file into the Word document in different ways.
        #
        # No return value.
        #
        # See also: SetHyperlink InsertCaption InsertImage InsertList InsertText

        if { $pasteFormat ne "" } {
            set tmpAppId [Cawt GetApplicationId $rangeId]
            set tmpDocId [Word OpenDocument $tmpAppId [file nativename $fileName] false]
            set tmpRangeId [Word GetStartRange $tmpDocId]
            $tmpRangeId WholeStory
            $tmpRangeId Copy

            $rangeId PasteAndFormat [Word GetEnum $pasteFormat]

            # Workaround: Select a small portion of text and copy it to clipboard
            # to avoid an alert message regarding lots of data in clipboard.
            # Setting DisplayAlerts to false does not help here.
            set dummyRange [Word CreateRange $tmpDocId 0 1]
            $dummyRange Copy
            Cawt Destroy $dummyRange

            Word Close $tmpDocId
            Cawt Destroy $tmpRangeId
            Cawt Destroy $tmpDocId
            Cawt Destroy $tmpAppId
        } else {
            # InsertFile(FileName, Range, ConfirmConversions, Link, Attachment)
            $rangeId InsertFile [file nativename $fileName] \
                                "" \
                                [Cawt TclBool false] \
                                [Cawt TclBool false] \
                                [Cawt TclBool false]
        }
    }

    proc InsertImage { rangeId imgFileName { linkToFile false } { saveWithDoc true } } {
        # Insert an image into a range of a document.
        #
        # rangeId     - Identifier of the text range.
        # imgFileName - File name of the image (as absolute path).
        # linkToFile  - Insert a link to the image file.
        # saveWithDoc - Embed the image into the document.
        #
        # The file name of the image must be an absolute pathname. Use a
        # construct like [file join [pwd] "myImage.gif"] to insert
        # images from the current directory.
        #
        # Return the identifier of the inserted image as an inline shape.
        #
        # See also: ScaleImage CropImage InsertFile InsertCaption InsertList InsertText

        if { ! $linkToFile && ! $saveWithDoc } { 
            error "InsertImage: linkToFile and saveWithDoc are both set to false."
        }

	set fileName [file nativename $imgFileName]
        set shapeId [$rangeId -with { InlineShapes } AddPicture $fileName \
                  [Cawt TclInt $linkToFile] \
                  [Cawt TclInt $saveWithDoc]]
        return $shapeId
    }

    proc ScaleImage { shapeId scaleWidth scaleHeight } {
        # Scale an image.
        #
        # shapeId     - Identifier of the image inline shape.
        # scaleWidth  - Horizontal scale factor.
        # scaleHeight - Vertical scale factor.
        #
        # The scale factors are floating point values. 1.0 means no scaling.
        #
        # No return value.
        #
        # See also: InsertImage CropImage

        $shapeId LockAspectRatio [Cawt TclInt false]
        $shapeId ScaleWidth  [expr { 100.0 * double($scaleWidth) }]
        $shapeId ScaleHeight [expr { 100.0 * double($scaleHeight) }]
    }

    proc CropImage { shapeId { cropBottom 0.0 } { cropTop 0.0 } { cropLeft 0.0 } { cropRight 0.0 } } {
        # Crop an image at the four borders.
        #
        # shapeId    - Identifier of the image inline shape.
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
        # See also: InsertImage ScaleImage ::Cawt::CentiMetersToPoints ::Cawt::InchesToPoints

        $shapeId -with { PictureFormat } CropBottom $cropBottom
        $shapeId -with { PictureFormat } CropTop    $cropTop
        $shapeId -with { PictureFormat } CropLeft   $cropLeft
        $shapeId -with { PictureFormat } CropRight  $cropRight
    }

    proc InsertCaption { rangeId labelId text { pos wdCaptionPositionBelow } } {
        # Insert a caption into a range of a document.
        #
        # rangeId - Identifier of the text range.
        # labelId - Value of enumeration type WdCaptionLabelID. 
        #           Possible values: wdCaptionEquation, wdCaptionFigure, wdCaptionTable
        # text    - Text of the caption.
        # pos     - Value of enumeration type WdCaptionPosition (see wordConst.tcl).
        #
        # Return the new extended range.
        #
        # See also: ConfigureCaption InsertFile InsertImage InsertList InsertText
 
        $rangeId InsertCaption [Word GetEnum $labelId] $text "" [Word GetEnum $pos] 0
        return $rangeId
    }

    proc ConfigureCaption { appId labelId chapterStyleLevel { includeChapterNumber true } \
                            { numberStyle wdCaptionNumberStyleArabic } \
                            { separator wdSeparatorHyphen } } {
        # Configure style of a caption type identified by its label identifier.
        #
        # appId                - Identifier of the Word instance.
        # labelId              - Value of enumeration type WdCaptionLabelID. 
        #                        Possible values: wdCaptionEquation, wdCaptionFigure, wdCaptionTable
        # chapterStyleLevel    - 1 corresponds to Heading1, 2 corresponds to Heading2, ...
        # includeChapterNumber - Flag indicating whether to include the chapter number.
        # numberStyle          - Value of enumeration type WdCaptionNumberStyle (see wordConst.tcl).
        # separator            - Value of enumeration type WdSeparatorType (see wordConst.tcl).
        #
        # No return value.
        #
        # See also: InsertCaption

        set captionItem [$appId -with { CaptionLabels } Item [Word GetEnum $labelId]]
        $captionItem ChapterStyleLevel    [expr $chapterStyleLevel]
        $captionItem IncludeChapterNumber [Cawt TclBool $includeChapterNumber]
        $captionItem NumberStyle          [Word GetEnum $numberStyle]
        $captionItem Separator            [Word GetEnum $separator]
    }

    proc AddTable { rangeId numRows numCols { spaceAfter -1 } } {
        # Add a new table in a text range.
        #
        # rangeId    - Identifier of the text range.
        # numRows    - Number of rows of the new table.
        # numCols    - Number of columns of the new table.
        # spaceAfter - Spacing in points after the table.
        #
        # Return the identifier of the new table.
        #
        # See also: GetNumRows GetNumColumns

        set docId [Word GetDocumentId $rangeId]
        set tableId [$docId -with { Tables } Add $rangeId $numRows $numCols]
        if { $spaceAfter >= 0 } {
            $tableId -with { Range ParagraphFormat } SpaceAfter $spaceAfter
        }
        Cawt Destroy $docId
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
        # Find a table by its index.
        #
        # docId - Identifier of the document.
        # index - Index of the table to find.
        #
        # Return the identifier of the found table.
        # If the index is out of bounds an error is thrown.
        #
        # See also: GetNumTables

        set count [Word GetNumTables $docId]

        if { $index < 1 || $index > $count } {
            error "GetTableIdByIndex: Invalid index $index given."
        }
        return [$docId -with { Tables } Item $index]
    }

    proc SetTableBorderLineStyle { tableId \
              { outsideLineStyle wdLineStyleSingle } \
              { insideLineStyle  wdLineStyleSingle } } {
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

        set border [$tableId Borders]
        $border OutsideLineStyle [Word GetEnum $outsideLineStyle]
        $border InsideLineStyle  [Word GetEnum $insideLineStyle]
        Cawt Destroy $border
    }

    proc SetTableBorderLineWidth { tableId \
              { outsideLineWidth wdLineWidth050pt } \
              { insideLineWidth  wdLineWidth050pt } } {
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

        set border [$tableId Borders]
        $border OutsideLineWidth [Word GetEnum $outsideLineWidth]
        $border InsideLineWidth  [Word GetEnum $insideLineWidth]
        Cawt Destroy $border
    }

    proc GetNumRows { tableId } {
        # Return the number of rows of a Word table.
        #
        # tableId - Identifier of the Word table.
        #
        # See also: GetNumColumns GetNumTables

        return [$tableId -with { Rows } Count]
    }

    proc GetNumColumns { tableId } {
        # Return the number of columns of a Word table.
        #
        # tableId - Identifier of the Word table.
        #
        # See also: GetNumRows GetNumTables

        return [$tableId -with { Columns } Count]
    }

    proc AddRow { tableId { beforeRowNum end } { numRows 1 } } {
        # Add one or more rows to a table.
        #
        # tableId      - Identifier of the Word table.
        # beforeRowNum - Insertion row number. Row numbering starts with 1.
        #                The new row is inserted before the given row number.
        #                If not specified or "end", the new row is appended at
        #                the end.
        # numRows      - Number of rows to be inserted.
        #
        # No return value.
        #
        # See also: GetNumRows

        Cawt PushComObjects

        set rowsId [$tableId Rows]
        if { $beforeRowNum eq "end" } {
            for { set r 1 } { $r <= $numRows } {incr r } {
                $rowsId Add
            }
        } else {
            if { $beforeRowNum < 1 || $beforeRowNum > [Word GetNumRows $tableId] } {
                error "AddRow: Invalid row number $beforeRowNum given."
            }
            set rowId [$tableId -with { Rows } Item $beforeRowNum]
            for { set r 1 } { $r <= $numRows } {incr r } {
                $rowsId Add $rowId
            }
        }

        Cawt PopComObjects
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
        Cawt Destroy $cellId
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
        Cawt Destroy $rowId
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
        Cawt Destroy $colId
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
        # No return value.
        #
        # See also: GetCellValue SetRowValues SetMatrixValues

        set rangeId [Word GetCellRange $tableId $row $col]
        $rangeId Text $val
        Cawt Destroy $rangeId
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

        set rangeId [Word GetCellRange $tableId $row $col]
        set val [Word TrimString [$rangeId Text]]
        Cawt Destroy $rangeId
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
            set len [Word GetNumColumns $tableId]
        } else {
            set len $numVals
        }
        set valList [list]
        set col $startCol
        set ind 0
        while { $ind < $len } {
            set val [Word GetCellValue $tableId $row $col]
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
        Cawt Destroy $colId
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
