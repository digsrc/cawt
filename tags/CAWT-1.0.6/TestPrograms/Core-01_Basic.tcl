# Test basic functionality of the CawtCore package.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"

# Load CAWT as a complete package and all sub-packages.
set retVal [catch {package require cawt} cawtVersion]
set retVal [catch {package require Img} imgVersion]
set retVal [catch {package require tablelist} tblVersion]
set retVal [catch {package require base64} base64Version]

puts [format "%-25s: %s" "Tcl version" [info patchlevel]]
puts [format "%-25s: %s" "Twapi version" [::Cawt::GetPkgVersion "twapi"]]
puts [format "%-25s: %s" "Img version" $imgVersion]
puts [format "%-25s: %s" "Tablelist version" $tblVersion]
puts [format "%-25s: %s" "Base64 version" $base64Version]
puts ""
puts [format "%-25s: %s" "CAWT version" $cawtVersion]
puts ""
puts [format "%-25s: %s" "CawtCore version"     [::Cawt::GetPkgVersion "cawtcore"]]
puts [format "%-25s: %s" "CawtEarth version"    [::Cawt::GetPkgVersion "cawtearth"]]
puts [format "%-25s: %s" "CawtExcel version"    [::Cawt::GetPkgVersion "cawtexcel"]]
puts [format "%-25s: %s" "CawtExplorer version" [::Cawt::GetPkgVersion "cawtexplorer"]]
puts [format "%-25s: %s" "CawtMatlab version"   [::Cawt::GetPkgVersion "cawtmatlab"]]
puts [format "%-25s: %s" "CawtOcr version"      [::Cawt::GetPkgVersion "cawtocr"]]
puts [format "%-25s: %s" "CawtPpt version"      [::Cawt::GetPkgVersion "cawtppt"]]
puts [format "%-25s: %s" "CawtWord version"     [::Cawt::GetPkgVersion "cawtword"]]

if { [lindex $argv 0] eq "auto" } {
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
