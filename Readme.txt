CAWT is a high-level Tcl interface for scripting Microsoft Windows® 
applications having a COM interface.
It uses Twapi for automation via the COM interface.
Currently packages for Microsoft Excel, Word, PowerPoint, Outlook and
Internet Explorer, MathWorks Matlab and Google Earth are available.

Note, that only Microsoft Office packages Excel, Word and PowerPoint
are in active developement.
The other packages are proof-of-concept examples only.

CAWT sources are available at https://sourceforge.net/projects/cawt/
The CAWT homepage is at http://www.poSoft.de/html/extCawt.html

The CAWT user distribution contains the Tcl sources, documentation (user 
and reference manual), several test programs showing the use of the CAWT
functionality and external libraries Twapi, TkImg, Base64 and Tablelist.

The CAWT developer distribution additionally contains scripts for 
generating the documentation, the distribution packages and the CAWT
Starkit. It also includes the external packages Ruff! and textutil.
The developer distribution is intended for programmers who want to
extend the CAWT package.

Release history:
================

2.0.0   2015-03-31
    Ensembled all CAWT namespaces.
    All Office enumerations are stored in module specific hash tables.
    Updated and extended user manual.
    Added application EnumExplorer.tcl to display Office enumerations.
    New module excelHtml.tcl for HTML export of Excel tables.
    External packages:
        Updated Twapi to version 4.1.27.
        Updated Img (32 and 64 bit) to version 1.4.3.
        Updated Tablelist to version 5.13.
    CawtExcel:
        New implementation of InsertImage.
    CawtWord:
        Extended procedure UpdateFields to update TablesOfContents and 
        TablesOfFigures of a document.
    New procedures in CawtCore:
        GetApplicationVersion, IsApplicationId, 
        PushComObjects, PopComObjects, 
        PrintNumComObjects, CheckComObjects, 
        GetComObjects, GetNumComObjects
        Replaced procedure IsValidId with IsComObject
    New procedures in CawtExcel:
        GetRangeAsIndex, GetRangeAsString, GetRangeTextColor
    New procedures in CawtWord:
        ScaleImage, SetInternalHyperlink, InsertFile, DiffWordFile

1.2.0   2014-12-14
    Compatibility issue: Incompatible changes in module CawtWord.
        Removed parameter docId from all procedures, which had both
        docId and rangeId parameters:
            SetRangeStartIndex, SetRangeEndIndex, ExtendRange,
            AddText, SetHyperlink, AddTable.
    CawtExcel: Added optional startRow parameter to TablelistToWorksheet.
    Extended test suite for changed and new procedures.
    New procedures in CawtWord:
        GetDocumentId, SetRangeFontUnderline, CreateRangeAfter,
        InsertCaption, ConfigureCaption,
        AddBookmark, GetBookmarkName, SetLinkToBookmark,
        GetListGalleryId, GetListTemplateId, InsertList

1.1.0   2014-08-30
    Compatibility issue: Incompatible changes in module CawtWord.
        Unified signatures of AddText, AppendText, 
        AddParagraph, AppendParagraph.
        Changed handling of text ranges.
    New module CawtOutlook to control Microsoft Outlook applications.
        Currently only functionality for creating and sending mails.
    Extended test suite for changed and new procedures.
    New procedures in CawtExcel:
        FreezePanes, ScreenUpdate
    New procedures in CawtWord:
        SelectRange, GetRangeInformation, CreateRange, SetRangeFontName, 
        SetRangeStyle, SetRangeFontSize,
        InsertText, AddText, GetNumCharacters,
        AddPageBreak, ToggleSpellCheck

1.0.7   2014-06-14
    Updated Twapi version to official 4.0.61.
    Extended test suite for changed and new procedures.
    CawtExcel:
        Added support for CSV files with multi-line cells.
    CawtPpt:
       Extended CopySlide to copy slides between presentations.
       Extended AddPres with optional parameter for template file.
       Extended AddSlide to supply custom layout object as type parameter.
    New procedures in CawtCore:
        ColorToRgb
    New procedures in CawtExcel:
        UseImgTransparency, WorksheetToImg, ImgToWorksheet,
        SetRowHeight, SetRowsHeight, GetRangeFillColor,
        SetHyperlinkToFile, SetHyperlinkToCell, SetLinkToCell,
        SetRangeTooltip
    New procedures in CawtPpt:
        MoveSlide, GetTemplateExtString, GetNumCustomLayouts,
        GetCustomLayoutName, GetCustomLayoutId

1.0.6   2014-04-21
    Improved and extended test suite.
    Updated Twapi version to 4.0b53 to fix a bug with sparse matrices as
    well as core dumps with Word 2013.
    Improved and corrected handling of sparse matrices in Excel.
    Bug fix in excelCsv module.
    Possible incompatibility in GetRowValues and GetColumnValues:
        Changed startRow resp. startCol to default value 0 instead of 1.
    New procedures in CawtExcel:
        GetWorksheetAsMatrix, GetMaxRows, GetMaxColumns, GetFirstUsedRow,
        GetLastUsedRow, GetFirstUsedColumn, GetLastUsedColumn.

1.0.5   2014-01-26
    New procedures in CawtExcel:
        SetCommentDisplayMode, SetRangeComment, SetRangeMergeCells, 
        SetRangeFontSubscript, SetRangeFontSuperscript, GetRangeCharacters.

1.0.4   2013-11-23
    Improved test suite.
    Added support for Office 2013.
    Added support for 64-bit Office.
    Updated Img extension to version 1.4.2 (32-bit and 64-bit).
    Update Tablelist to version 5.10.
    New procedures in CawtWord:
        SaveAsPdf, UpdateFields, CropImage.
    New procedures in CawtExcel:
        CopyWorksheetBefore, CopyWorksheetAfter, 
        GetWorksheetIndexByName, IsWorksheetProtected, 
        IsWorksheetVisible, SetWorksheetTabColor,
        UnhideWorksheet, DiffExcelFiles.

1.0.3   2013-08-30
    New procedures in CawtExcel:
        ExcelFileToMediaWikiFile, ExcelFileToWikitFile,
        ExcelFileToRawImageFile, RawImageFileToExcelFile,
        ExcelFileToMatlabFile, MatlabFileToExcelFile,
        GetTablelistValues, SetTablelistValues.

1.0.2   2013-07-28
    Updated Twapi version to 4.0b22.
    Updated Img version to 1.4.1.
    Added new module CawtOcr. 
    New procedures in CawtCore:
        Clipboard2Img, Img2Clipboard
    New procedures in CawtExcel: 
        SetRangeBorder 

1.0.1   2013-04-28
    Extended Excel chart generation. 
    Updated Twapi version to 4.0a16. 
    Added support to generate a CAWT starkit.

1.0.0   2012-12-23
    Replaced Tcom with Twapi for COM access.
    Added support for PowerPoint, InternetExplorer, GoogleEarth and Matlab.
    Added user and reference manual.
    Unification of procedure names.
    Support for Microsoft Office versions 2003, 2007, 2010.
