# Test basic functionality of the CawtMatlab package.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}
set retVal [catch {package require cawt} pkgVersion]

set appId [Matlab OpenNew]

puts [format "%-25s: %s" "Tcl version" [info patchlevel]]
puts [format "%-25s: %s" "Cawt version" $pkgVersion]
puts [format "%-25s: %s" "Twapi version" [Cawt GetPkgVersion "twapi"]]

puts [format "%-25s: %s" "Active Printer" \
                        [Cawt GetActivePrinter $appId]]

puts [format "%-25s: %s" "User Name" \
                        [Cawt GetUserName $appId]]

puts [format "%-25s: %s" "Startup Pathname" \
                         [Cawt GetStartupPath $appId]]
puts [format "%-25s: %s" "Templates Pathname" \
                         [Cawt GetTemplatesPath $appId]]
puts [format "%-25s: %s" "Add-ins Pathname" \
                         [Cawt GetUserLibraryPath $appId]]
puts [format "%-25s: %s" "Installation Pathname" \
                         [Cawt GetInstallationPath $appId]]
puts [format "%-25s: %s" "User Folder Pathname" \
                         [Cawt GetUserPath $appId]]

if { [lindex $argv 0] eq "auto" } {
    Matlab Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
