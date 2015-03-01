# Test basic functionality of the CawtExplorer package.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
set retVal [catch {package require cawt} pkgVersion]

set appId [Explorer OpenNew false]

puts [format "%-30s: %s" "Tcl version" [info patchlevel]]
puts [format "%-30s: %s" "Cawt version" $pkgVersion]
puts [format "%-30s: %s" "Twapi version" [::Cawt::GetPkgVersion "twapi"]]

puts [format "%-30s: %s" "Active Printer" \
                        [::Cawt::GetActivePrinter $appId]]

puts [format "%-30s: %s" "User Name" \
                        [::Cawt::GetUserName $appId]]

puts [format "%-30s: %s" "Startup Pathname" \
                         [::Cawt::GetStartupPath $appId]]
puts [format "%-30s: %s" "Templates Pathname" \
                         [::Cawt::GetTemplatesPath $appId]]
puts [format "%-30s: %s" "Add-ins Pathname" \
                         [::Cawt::GetUserLibraryPath $appId]]
puts [format "%-30s: %s" "Installation Pathname" \
                         [::Cawt::GetInstallationPath $appId]]
puts [format "%-30s: %s" "User Folder Pathname" \
                         [::Cawt::GetUserPath $appId]]

puts [format "%-30s: %s" "Appl. name (from Application)" \
         [::Cawt::GetApplicationName $appId]]

if { [lindex $argv 0] eq "auto" } {
    Explorer Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
