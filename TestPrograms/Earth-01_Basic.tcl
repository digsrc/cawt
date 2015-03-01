# Test basic functionality of the CawtEarth package.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
set retVal [catch {package require cawt} pkgVersion]

set appId [Earth OpenNew]

puts [format "%-25s: %s" "Tcl version" [info patchlevel]]
puts [format "%-25s: %s" "Cawt version" $pkgVersion]
puts [format "%-25s: %s" "Twapi version" [::Cawt::GetPkgVersion "twapi"]]

puts [format "%-25s: %s.%s.%s (%s)" "Google Earth Version" \
                             [$appId versionMajor] \
                             [$appId versionMinor] \
                             [$appId VersionBuild] \
                             [$appId VersionAppType]]

if { [lindex $argv 0] eq "auto" } {
    Earth Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
