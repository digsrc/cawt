# Test basic functionality of the CawtOcr package.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
set retVal [catch {package require cawt} pkgVersion]

set appId [::Ocr::Open]

puts [format "%-25s: %s" "Tcl version" [info patchlevel]]
puts [format "%-25s: %s" "Cawt version" $pkgVersion]
puts [format "%-25s: %s" "Twapi version" [::Cawt::GetPkgVersion "twapi"]]

if { [lindex $argv 0] eq "auto" } {
    ::Ocr::Close $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
