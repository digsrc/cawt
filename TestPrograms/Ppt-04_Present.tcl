# Test CawtPpt procedures for presenting a slide show.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

set appId [Ppt Open]

set pptFile [file join [pwd] "testOut" "Ppt-02_Misc"]
append pptFile [Ppt GetExtString $appId]

set presId [Ppt OpenPres $appId $pptFile true]

Ppt ShowSlide $presId 2
after 1000

set slideCount [Ppt GetNumSlides $presId]
puts "Have $slideCount slides"

puts "Have [Ppt GetNumSlideShows $appId] SlideShows"
set slideShowId [Ppt UseSlideShow $presId 1]
puts "Have [Ppt GetNumSlideShows $appId] SlideShows"

for { set i 0 } { $i < 3 } { incr i } {
    for { set s 1 } { $s < $slideCount } { incr s } {
        Ppt SlideShowNext $slideShowId
        after 500
    }
    Ppt SlideShowFirst $slideShowId
    after 500
}
after 500
Ppt SlideShowLast $slideShowId
after 500
Ppt SlideShowPrev $slideShowId

Ppt ExitSlideShow $slideShowId

# TODO
# With ActivePresentation.SlideShowSettings
#        .ShowType = ppShowTypeSpeaker
#        .LoopUntilStopped = true
#        .ShowWithNarration = msoTrue
#        .ShowWithAnimation = msoTrue
#        .RangeType = ppShowAll
#        .AdvanceMode = ppSlideShowUseSlideTimings
#        .PointerColor.RGB = RGB(Red:=255, Green:=0, Blue:=0)
#        .Run
#    End With

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Ppt Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
