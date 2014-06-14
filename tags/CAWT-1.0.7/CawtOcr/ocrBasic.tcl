# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.
# Idea taken from http://wiki.tcl.tk/38305

namespace eval ::Ocr {

    variable ocrAppName "MODI.Document"
    variable _ruffdoc

    lappend _ruffdoc Introduction {
        The Ocr namespace provides commands to control Microsoft Document Imaging.
    }

    # Hierarchy: [Document [Images [Layout [Text]]]

    proc Open {} {
        # Open an OCR document instance.
        #
        # Return the OCR document identifier.
        #
        # See also: OpenDocument Close

	variable ocrAppName

        set docId [::Cawt::GetOrCreateApp $ocrAppName true]
        return $docId
    }

    proc OpenDocument { docId fileName } {
        # Open an image file for OCR scanning.
        #
        # docId    - OCR document identifier.
        # fileName - Image to be scanned. Must be in TIFF or BMP format.
        #
        # No return value.
        #
        # See also: Open Close

        $docId Create $fileName
    }

    proc Close { docId } {
        # Close an OCR document instance.
        #
        # docId - Identifier of the OCR document.
        #
        # No return value.
        #
        # See also: Open

        $docId Close
    }

    proc GetNumImages { docId } {
        # Return the number of images of an OCR document.
        #
        # docId - Identifier of the OCR document.
        #
        # See also: OpenDocument Scan

        return [$docId -with { Images } Count]
    }

    proc Scan { docId { imgNum 0 } } {
        # Scan an image.
        #
        # docId  - Identifier of the OCR document.
        # imgNum - Image number to be scanned.
        #
        # Return the layout identifier of the scanned image.
        #
        # See also: OpenDocument GetNumImages

        $docId OCR
        set imgId [$docId -with { Images } Item [expr int($imgNum)]]
        set layoutId [$imgId Layout]
        ::Cawt::Destroy $imgId
        return $layoutId
    }

    proc GetFullText { layoutId } {
        # Return the recognized text of a OCR layout.
        #
        # layoutId - Identifier of the OCR layout.
        #
        # See also: Scan

        return [$layoutId Text]
    }

    proc GetNumWords { layoutId } {
        # Return the number of words identified in an OCR document.
        #
        # layoutId - Identifier of the OCR layout.
        #
        # See also: GetFullText GetNumImages Scan

        return [$layoutId -with { Words } Count]
    }

    proc GetWord { layoutId wordNum } {
        # Return the text of a recognized word.
        #
        # layoutId - Identifier of the OCR layout.
        # wordNum  - Index number of the word (starting at zero).
        #
        # See also: GetFullText GetNumWords Scan

        set word [$layoutId -with { Words } Item [expr int($wordNum)]]
        set wordText [$word Text]
        ::Cawt::Destroy $word
        return $wordText
    }

    proc GetWordStats { layoutId wordNum } {
        # Return statistics of a recognized word.
        #
        # layoutId - Identifier of the OCR layout.
        # wordNum  - Index number of the word (starting at zero).
        #
        # The statistics is returned as a dictionary containing the
        # following keys: Id, LineId, RegionId, FontId, Confidence.
        #
        # See also: GetFullText GetWord Scan

        set word [$layoutId -with { Words } Item [expr int($wordNum)]]
        dict set wordStats "Id" [$word Id]
        dict set wordStats "LineId" [$word LineId]
        dict set wordStats "RegionId" [$word RegionId]
        dict set wordStats "FontId" [$word FontId]
        dict set wordStats "Confidence" [$word RecognitionConfidence]
        ::Cawt::Destroy $word
        return $wordStats
    }
}
