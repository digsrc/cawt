# Test basic functionality of the CawtCore package.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}

# Load CAWT as a complete package and all sub-packages.
set retVal [catch {package require cawt} cawtVersion]
set retVal [catch {package require Img} imgVersion]
set retVal [catch {package require tablelist} tblVersion]
set retVal [catch {package require base64} base64Version]

puts [format "%-25s: %s" "Tcl version" [info patchlevel]]
puts [format "%-25s: %s" "Twapi version" [Cawt GetPkgVersion "twapi"]]
puts [format "%-25s: %s" "Img version" $imgVersion]
puts [format "%-25s: %s" "Tablelist version" $tblVersion]
puts [format "%-25s: %s" "Base64 version" $base64Version]
puts ""
puts [format "%-25s: %s" "CAWT version" $cawtVersion]
puts ""
puts [format "%-25s: %s" "CawtCore version"     [Cawt GetPkgVersion "cawtcore"]]
puts [format "%-25s: %s" "CawtEarth version"    [Cawt GetPkgVersion "cawtearth"]]
puts [format "%-25s: %s" "CawtExcel version"    [Cawt GetPkgVersion "cawtexcel"]]
puts [format "%-25s: %s" "CawtExplorer version" [Cawt GetPkgVersion "cawtexplorer"]]
puts [format "%-25s: %s" "CawtMatlab version"   [Cawt GetPkgVersion "cawtmatlab"]]
puts [format "%-25s: %s" "CawtOcr version"      [Cawt GetPkgVersion "cawtocr"]]
puts [format "%-25s: %s" "CawtPpt version"      [Cawt GetPkgVersion "cawtppt"]]
puts [format "%-25s: %s" "CawtWord version"     [Cawt GetPkgVersion "cawtword"]]
puts ""
set r [Cawt RgbToColor 255 0 0]
set g [Cawt RgbToColor 0 255 0]
set b [Cawt RgbToColor 0 0 255]
puts [format "Red Green Blue as Office color: %08X %08X %08X" $r $g $b]
puts "Red Green Blue as RGB color: \
     [Cawt ColorToRgb $r] \
     [Cawt ColorToRgb $g] \
     [Cawt ColorToRgb $b]"

if { [lindex $argv 0] eq "full" } {
    puts "Testing color conversion procedures (both directions for all r g b values) ..."
    for { set r 0 } { $r < 256 } { incr r } {
        for { set g 0 } { $g < 256 } { incr g } {
            for { set b 0 } { $b < 256 } { incr b } {
                set colorNum [Cawt RgbToColor $r $g $b]
                set rgb [Cawt ColorToRgb $colorNum]
                Cawt CheckNumber $r [lindex $rgb 0] "Convert color $r $g $b" false
                Cawt CheckNumber $g [lindex $rgb 1] "Convert color $r $g $b" false
                Cawt CheckNumber $b [lindex $rgb 2] "Convert color $r $g $b" false
            }
        }
    }
    puts "Conversion test finished."
}

Cawt CheckNumber 72.0 [Cawt InchesToPoints 1]  "InchesToPoints"
Cawt CheckNumber  1.0 [Cawt PointsToInches 72] "PointsToInches"

Cawt CheckNumber 1.0 [Cawt PointsToInches      [Cawt InchesToPoints      1]] "InchesToPoints"
Cawt CheckNumber 1.0 [Cawt PointsToCentiMeters [Cawt CentiMetersToPoints 1]] "CentiMetersToPoints"

if { [lindex $argv 0] eq "auto" } {
    Cawt Destroy
    exit 0
}
Cawt Destroy
