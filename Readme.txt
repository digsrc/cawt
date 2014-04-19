CAWT is a high-level Tcl interface for scripting Microsoft Windows® applications
having a COM interface. It uses Twapi for automation via the COM interface.
Currently packages for Microsoft Excel, Word, PowerPoint and Internet Explorer,
MathWorks Matlab and Google Earth are available.

CAWT sources are available at https://sourceforge.net/projects/cawt/
The CAWT homepage is at http://www.poSoft.de/html/extCawt.html

The CAWT user distribution contains the Tcl sources, documentation (user and 
reference manual), several test programs showing the use of the CAWT functionality
and the needed external libraries Twapi, TkImg, Base64 and Tablelist.

The CAWT developer distribution additionally contains scripts for generating the 
documentation,  the distribution packages and the CAWT Starkit. It also includes 
the external packages Ruff! and textutil. The developer distribution is intended
for programmers who want to extend the CAWT package.

Release history:
================

1.0.0	2012-12-23
	Replaced Tcom with Twapi for COM access.
        Added support for PowerPoint, Internet Explorer, Google Earth and Matlab.
        Added user and reference manual.
        Unification of procedure names.
        Support for Microsoft Office versions 2003, 2007, 2010.

1.0.1	2013-04-28
	Extended Excel chart generation. 
        Updated Twapi version to 4.0a16. 
        Added support to generate a CAWT starkit.

1.0.2	2013-07-28
	Updated Twapi version to 4.0b22.
        Updated Img version to 1.4.1.
        Added new module CawtOcr. 
        New procedures: SetRangeBorder, Clipboard2Img, Img2Clipboard.

1.0.3	2013-08-30
	New procedures in module CawtExcel:
            ExcelFileToMediaWikiFile, ExcelFileToWikitFile, ExcelFileToRawImageFile,
            RawImageFileToExcelFile, ExcelFileToMatlabFile, MatlabFileToExcelFile,
            GetTablelistValues, SetTablelistValues.

1.0.4	2013-11-23
	Improved test suite.
        Added support for Office 2013.
        Added support for 64-bit Office.
        Updated Img extension to version 1.4.2 (32-bit and 64-bit).
        Update Tablelist to version 5.10.
        New procedures in CawtWord:
            SaveAsPdf, UpdateFields, CropImage.
        New procedures in CawtExcel:
            CopyWorksheetBefore, CopyWorksheetAfter, 
            GetWorksheetIndexByName, IsWorksheetProtected, IsWorksheetVisible,
            SetWorksheetTabColor, UnhideWorksheet, DiffExcelFiles.

1.0.5	2014-01-26
	New procedures in CawtExcel:
            SetCommentDisplayMode, SetRangeComment, SetRangeMergeCells, 
            SetRangeFontSubscript, SetRangeFontSuperscript, GetRangeCharacters.

1.0.6	2014-04-21
	Improved and extended test suite.
        Updated Twapi version to 4.0b53 to fix a bug with sparse matrices as well
        as core dumps with Word 2013.
        Improved and corrected handling of sparse matrices in Excel.
        Bug fix in excelCsv module.
        Possible incompatibility in GetRowValues and GetColumnValues:
            Changed startRow resp. startCol to default value 0 instead of 1.
        New procedures in CawtExcel:
            GetWorksheetAsMatrix, GetMaxRows, GetMaxColumns, GetFirstUsedRow,
            GetLastUsedRow, GetFirstUsedColumn, GetLastUsedColumn.

