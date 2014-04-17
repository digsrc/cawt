# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval ::Ppt {

    variable pptVersion  "0.0"
    variable pptAppName  "PowerPoint.Application"
    variable _ruffdoc

    lappend _ruffdoc Introduction {
        The Ppt namespace provides commands to control Microsoft PowerPoint.
    }

    proc GetVersion { appId { useString false } } {
        # Return the version of a PowerPoint application.
        #
        # appId     - Identifier of the Excel instance.
        # useString - true: Return the version name (ex. "PowerPoint 2003").
        #             false: Return the version number (ex. "11.0").
        #
        # Both version name and version number are returned as strings.
        # Version number is in a format, so that it can be evaluated as a
        # floating point number.

        array set map {
            "8.0"  "PowerPoint 97"
            "9.0"  "PowerPoint 2000"
            "10.0" "PowerPoint 2002"
            "11.0" "PowerPoint 2003"
            "12.0" "PowerPoint 2007"
            "14.0" "PowerPoint 2010"
            "15.0" "PowerPoint 2013"
        }
        set version [$appId Version]
        if { $useString } {
            if { [info exists map($version)] } {
                return $map($version)
            } else {
                return "Unknown PowerPoint version"
            }
        } else {
            return $version
        }
        return $version
    }

    proc GetExtString { { appId "" } } {
        # Return the default extension of a PowerPoint file.
        #
        # appId - Identifier of the PowerPoint instance.
        #         OBSOLETE: This identifier is not needed anymore
        #         and will be removed in next major release.
        #
        # Starting with PowerPoint 12 (2007) this is the string ".pptx".
        # In previous versions it was ".ppt".

        variable pptVersion

        if { $pptVersion >= 12.0 } {
            return ".pptx"
        } else {
            return ".ppt"
        }
    }

    proc OpenNew { { width -1 } { height -1 } } {
        # Open a new PowerPoint instance.
        #
        # width   - Width of the application window. If negative, open with last used width.
        # height  - Height of the application window. If negative, open with last used height.
        #
        # Return the identifier of the new PowerPoint application instance.
        #
        # See also: Open Quit

	variable pptAppName
	variable pptVersion

        set appId [::Cawt::GetOrCreateApp $pptAppName false]
        set pptVersion [::Ppt::GetVersion $appId]
        ::Ppt::Visible $appId true
        if { $width >= 0 } {
            $appId Width [expr $width]
        }
        if { $height >= 0 } {
            $appId Height [expr $height]
        }
        return $appId
    }

    proc Open { { width -1 } { height -1 } } {
        # Open a PowerPoint instance. Use an already running PowerPoint, if available.
        #
        # width   - Width of the application window. If negative, open with last used width.
        # height  - Height of the application window. If negative, open with last used height.
        #
        # Return the identifier of the PowerPoint application instance.
        #
        # See also: OpenNew Quit

	variable pptAppName
	variable pptVersion

        set appId [::Cawt::GetOrCreateApp $pptAppName true]
        set pptVersion [::Ppt::GetVersion $appId]
        ::Ppt::Visible $appId true
        if { $width >= 0 } {
            $appId Width [expr $width]
        }
        if { $height >= 0 } {
            $appId Height [expr $height]
        }
        return $appId
    }

    proc Quit { appId { showAlert true } } {
        # Quit a PowerPoint application instance.
        #
        # appId     - Identifier of the PowerPoint instance.
        # showAlert - true: Show an alert window, if there are unsaved changes.
        #             false: Quit without saving any changes.
        #
        # No return value.
        #
        # See also: Open

        if { ! $showAlert } {
            ::Cawt::ShowAlerts $appId false
        }
        $appId Quit
    }

    proc Visible { appId visible } {
        # Toggle the visibilty of a PowerPoint application window.
        #
        # appId   - Identifier of the PowerPoint instance.
        # visible - true: Show the application window.
        #           false: Hide the application window.
        #
        # No return value.
        #
        # See also: Open OpenNew

        $appId Visible [::Cawt::TclInt $visible]
    }


    proc Close { presId } {
        # Close a presentation without saving changes.
        #
        # presId - Identifier of the presentation to close.
        #
        # Use the SaveAs method before closing, if you want to save changes.
        #
        # No return value.
        #
        # See also: SaveAs CloseAll

        $presId Close
    }

    proc CloseAll { appId } {
        # Close all presentations of a PowerPoint instance.
        #
        # appId - Identifier of the PowerPoint instance.
        #
        # Use the SaveAs method before closing, if you want to save changes.
        #
        # No return value.
        #
        # See also: SaveAs Close

        set numWins [$appId -with { Windows } Count]
        for { set ind $numWins } { $ind >= 1 } { incr ind -1 } {
            [$appId -with { Windows } Item $ind] Activate
            $appId -with { ActiveWindow } Close
        }
    }

    proc SaveAs { presId fileName { fmt "" } { embedFonts true } } {
        # Save a presentation to a PowerPoint file.
        #
        # presId     - Identifier of the presentation to save.
        # fileName   - Name of the PowerPoint file.
        # fmt        - Value of enumeration type PpSaveAsFileType (see pptConst.tcl).
        #              If not given or the empty string, the file is stored in the native
        #              format corresponding to the used PowerPoint version (ppSaveAsDefault).
        # embedFonts - true: Embed TrueType fonts.
        #              false: Do not embed TrueType fonts.
        #
        # Note: If "fmt" is not a PowerPoint format, but an image format, PowerPoint takes the
        #       specified file name and creates a directory with that name. Then it copies all
        #       slides as images into that directory. The slide images are automatically named by
        #       PowerPoint (ex. in German versions the slides are called Folie1.gif, Folie2.gif, ...).
        #       Use the ExportSlide procedure, if you want full control over image file names.
        #
        # No return value.
        #
        # See also: ExportSlides ExportSlide

        set fileName [file nativename $fileName]
        set appId [::Cawt::GetApplicationId $presId]
        ::Cawt::ShowAlerts $appId false
        if { $fmt eq "" } {
            $presId SaveAs $fileName
        } else {
            $presId -callnamedargs SaveAs \
                     FileName $fileName \
                     FileFormat [expr $fmt] \
                     EmbedTrueTypeFonts [::Cawt::TclInt $embedFonts]
        }
        ::Cawt::ShowAlerts $appId true
        ::Cawt::Destroy $appId
    }

    proc AddPres { appId  } {
        # Add a new empty presentation.
        #
        # appId - Identifier of the PowerPoint instance.
        #
        # Return the identifier of the new presentation.
        #
        # See also: OpenPres GetActivePres

        set presId [$appId -with { Presentations } Add]
        return $presId
    }

    proc OpenPres { appId fileName { readOnly false } } {
        # Open a presentation, i.e load a PowerPoint file.
        #
        # appId    - Identifier of the PowerPoint instance.
        # fileName - Name of the PowerPoint file (as absolute path).
        # readOnly - true: Open the presentation in read-only mode.
        #            false: Open the presentation in read-write mode.
        #
        # Return the identifier of the opened presentation. If the presentation was already
        # open, activate that presentation and return the identifier to that presentation.
        #
        # See also: AddPres GetActivePres

        set nativeName  [file nativename $fileName]
        set presentations [$appId Presentations]
        set retVal [catch {[$presentations Item [file tail $fileName]] Activate} d]
        if { $retVal == 0 } {
            puts "$nativeName already open"
            set presId [$presentations Item [file tail $fileName]]
        } else {
            set presId [$presentations Open $nativeName [::Cawt::TclInt $readOnly]]
        }
        ::Cawt::Destroy $presentations
        return $presId
    }

    proc GetActivePres { appId } {
        # Return the active presentation of an application.
        #
        # appId - Identifier of the PowerPoint instance.
        #
        # Return the identifier of the active presentation.
        #
        # See also: OpenPres AddPres

        return [$appId ActivePresentation]
    }

    proc SetViewType { presId viewType } {
        # Set the view type of a presentation.
        #
        # presId   - Identifier of the presentation.
        # viewType - Value of enumeration type PpViewType (see pptConst.tcl).
        #
        # No return value.
        #
        # See also: GetViewType

        set appId [::Cawt::GetApplicationId $presId]
        set actWin [$appId ActiveWindow]
        $actWin ViewType $viewType
        ::Cawt::Destroy $actWin
        ::Cawt::Destroy $appId
    }

    proc GetViewType { presId } {
        # Return the view type of a presentation.
        #
        # presId - Identifier of the presentation.
        #
        # See also: SetViewType

        set appId [::Cawt::GetApplicationId $presId]
        set actWin [$appId ActiveWindow]
        set viewType [$actWin ViewType]
        ::Cawt::Destroy $actWin
        ::Cawt::Destroy $appId
        return $viewType
    }

    proc AddSlide { presId { type $::Ppt::ppLayoutBlank } { slideIndex -1 } } {
        # Add a new slide to a presentation.
        #
        # presId     - Identifier of the presentation.
        # type       - Value of enumeration type PpSlideLayout (see pptConst.tcl).
        # slideIndex - Insertion index of new slide. If negative, insert slide at the end.
        #
        # Return the identifier of the new slide.
        #
        # See also: CopySlide

        if { $slideIndex < 0 } {
            set insertIndex [expr [::Ppt::GetNumSlides $presId] +1]
        } else {
            set insertIndex $slideIndex
        }
        set newSlide [$presId -with { Slides } Add $insertIndex [expr $type]]
        set newSlideIndex [::Ppt::GetSlideIndex $newSlide]
        ::Ppt::ShowSlide $presId $newSlideIndex
        return $newSlide
    }

    proc CopySlide { presId fromSlideIndex { toSlideIndex -1 } } {
        # Copy the contents of a slide into another slide.
        #
        # presId         - Identifier of the presentation.
        # fromSlideIndex - Index of source slide.
        # toSlideIndex   - Insertion index of copied slide. If negative, insert slide at the end.
        #
        # A new empty slide is created at the insertion index and the contents of the source
        # slide are copied into the new slide.
        #
        # Return the identifier of the new slide.
        #
        # See also: AddSlide

        if { $toSlideIndex < 0 } {
            set insertIndex [expr [::Ppt::GetNumSlides $presId] +1]
        } else {
            set insertIndex $toSlideIndex
        }
        set newSlide [::Ppt::AddSlide $presId $::Ppt::ppLayoutBlank $insertIndex]
        set newSlideIndex [::Ppt::GetSlideIndex $newSlide]

        set appId [::Cawt::GetApplicationId $presId]
        [$presId -with { Windows } Item 1] Activate
        set actWin [$appId ActiveWindow]

        # Save current view type.
        set saveViewType [::Ppt::GetViewType $presId]

        # Select slide to be copied.
        $actWin ViewType $::Ppt::ppViewSlide
        ::Ppt::ShowSlide $presId $fromSlideIndex

        set actSel [$actWin Selection]
        $actSel -with { SlideRange Shapes } SelectAll
        $actWin -with { Selection } Copy

        ::Ppt::ShowSlide $presId $newSlideIndex
        $appId -with { ActiveWindow View } Paste

        # Restore original view type.
        ::Ppt::SetViewType $presId $saveViewType

        ::Cawt::Destroy $actSel
        ::Cawt::Destroy $actWin
        ::Cawt::Destroy $appId

        return $newSlide
    }

    proc ExportSlide { slideId outputFile { imgType "GIF" } { width -1 } { height -1 } } {
        # Export a slide as an image.
        #
        # slideId    - Identifier of the slide.
        # outputFile - Name of the output file (as absolute path).
        # imgType    - Name of the image format filter. This is the name as stored in
        #              the Windows registry. Ex: "GIF", "PNG".
        # width      - Width of the generated images in pixels.
        # height     - Height of the generated images in pixels.
        #
        # If width and height are not specified or less than zero, the default sizes
        # of PowerPoint are used.
        #
        # No return value. If the export failed, an error is thrown.
        #
        # See also: ExportPptFile ExportSlides

        set nativeName [file nativename $outputFile]
        if { $width >= 0 && $height >= 0 } {
            set retVal [catch {$slideId Export $nativeName $imgType $width $height} errMsg]
        } else {
            set retVal [catch {$slideId Export $nativeName $imgType} errMsg]
        }
        if { $retVal } {
            error "Slide export failed. ( $errMsg )"
        }
    }

    proc ExportSlides { presId outputDir outputFileFmt { startIndex 1 } { endIndex "end" } \
                        { imgType "GIF" } { width -1 } { height -1 } } {
        # Export a range of slides as image files.
        #
        # presId        - Identifier of the presentation.
        # outputDir     - Name of the output folder (as absolute path).
        # outputFileFmt - Name of the output file names (C printf style with one "%d" for the slide index).
        # startIndex    - Start index for slide export.
        # endIndex      - End index for slide export.
        # imgType       - Name of the image format filter. This is the name as stored in
        #                 the Windows registry. Ex: "GIF", "PNG".
        # width         - Width of the generated images in pixels.
        # height        - Height of the generated images in pixels.
        #
        # If the output directory does not exist, it is created.
        # If width and height are not specified or less than zero, the default sizes
        # of PowerPoint are used.
        #
        # No return value. If the export failed, an error is thrown.
        #
        # See also: ExportPptFile ExportSlide

        set numSlides [::Ppt::GetNumSlides $presId]
        if { $startIndex < 1 || $startIndex > $numSlides } {
            error "startIndex ($startIndex) not in slide range."
        }
        if { $endIndex eq "end" } {
            set endIndex $numSlides
        }
        if { $endIndex < 1 || $endIndex > $numSlides || $endIndex < $startIndex } {
            error "endIndex ($endIndex) not in slide range."
        }

        if { ! [file isdir $outputDir] } {
            file mkdir $outputDir
        }
        set nativeName [file nativename $outputDir]

        for { set i $startIndex } { $i <= $endIndex } { incr i } {
            set slideId [::Ppt::GetSlideId $presId $i]
            set outputFile [format [file join $outputDir $outputFileFmt] $i]
            ::Ppt::ExportSlide $slideId $outputFile $imgType $width $height
            ::Cawt::Destroy $slideId
        }
    }

    proc ShowSlide { presId slideIndex } {
        # Show a specific slide.
        #
        # presId     - Identifier of the presentation.
        # slideIndex - Index of slide. Slide indices start at 1.
        #
        # No return value.

        set slideId [$presId -with { Slides } Item $slideIndex]
        $slideId Select
        ::Cawt::Destroy $slideId
    }

    proc GetNumSlides { presId } {
        # Return the number of slides of a presentation.
        #
        # presId - Identifier of the presentation.
        #
        # See also: GetNumSlideShows

        return [$presId -with { Slides } Count]
    }

    proc GetSlideIndex { slideId } {
        # Return the index of a slide.
        #
        # slideId - Identifier of the slide.
        #
        # See also: GetNumSlides AddSlide

        return [$slideId SlideIndex]
    }

    proc GetSlideId { presId slideIndex } {
        # Get slide identifier form slide index.
        #
        # presId     - Identifier of the presentation.
        # slideIndex - Index of slide. Slide indices start at 1.
        #
        # Return the identifier of the slide.

        set slideId [$presId -with { Slides } Item $slideIndex]
        return $slideId
    }

    proc GetNumSlideShows { appId } {
        # Return the number of slide shows of a presentation.
        #
        # appId - Identifier of the PowerPoint instance.
        #
        # See also: GetNumSlides UseSlideShow ExitSlideShow

        return [$appId -with { SlideShowWindows } Count]
    }

    proc UseSlideShow { presId slideShowIndex } {
        # Use specified slide show.
        #
        # presId         - Identifier of the presentation.
        # slideShowIndex - Index of the slide show. Indices start at 1.
        #
        # Return the identifier of the specified slide show.
        #
        # See also: GetNumSlides ExitSlideShow SlideShowNext

        $presId -with { SlideShowSettings } Run
        set appId [::Cawt::GetApplicationId $presId]
        set slideShow [$appId -with { SlideShowWindows } Item $slideShowIndex]
        ::Cawt::Destroy $appId
        return $slideShow
    }

    proc ExitSlideShow { slideShowId } {
        # Exit specified slide show.
        #
        # slideShowId - Identifier of the slide show as returned by UseSlideShow.
        #
        # No return value.
        #
        # See also: GetNumSlideShows UseSlideShow SlideShowNext

        $slideShowId -with { View } Exit
    }

    proc SlideShowNext { slideShowId } {
        # Goto next slide in slide show.
        #
        # slideShowId - Identifier of the slide show.
        #
        # No return value.
        #
        # See also: UseSlideShow SlideShowPrev SlideShowFirst SlideShowLast

        $slideShowId -with { View } Next
    }

    proc SlideShowPrev { slideShowId } {
        # Goto previous slide in slide show.
        #
        # slideShowId - Identifier of the slide show.
        #
        # No return value.
        #
        # See also: UseSlideShow SlideShowNext SlideShowFirst SlideShowLast

        $slideShowId -with { View } Previous
    }

    proc SlideShowFirst { slideShowId } {
        # Goto first slide in slide show.
        #
        # slideShowId - Identifier of the slide show.
        #
        # No return value.
        #
        # See also: UseSlideShow SlideShowNext SlideShowPrev SlideShowLast

        $slideShowId -with { View } First
    }

    proc SlideShowLast { slideShowId } {
        # Goto last slide in slide show.
        #
        # slideShowId - Identifier of the slide show.
        #
        # No return value.
        #
        # See also: UseSlideShow SlideShowNext SlideShowPrev SlideShowFirst

        $slideShowId -with { View } Last
    }

    proc InsertImage { slideId imgFileName left top { width -1 } { height -1 } } {
        # Insert an image into a slide.
        #
        # slideId     - Identifier of the slide where the image is inserted.
        # imgFileName - File name of the image (as absolute path).
        # left        - X position of top-left image position in points.
        # top         - Y position of top-left image position in points.
        # width       - Width of image in points.
        # height      - Height of image in points.
        #
        # The file name of the image must be an absolute pathname. Use a
        # construct like [file join [pwd] "myImage.gif"] to insert
        # images from the current directory.
        #
        # Return the identifier of the inserted image.
        #
        # See also: ::Cawt::InchesToPoints ::Cawt::CentiMetersToPoints

	set fileName [file nativename $imgFileName]
        set imgId [$slideId -with { Shapes } AddPicture $fileName \
                   [::Cawt::TclInt 0] [::Cawt::TclInt 1] \
                   $left $top $width $height]
        return $imgId
    }
}
