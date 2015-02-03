# Test basic functionality of the CawtPpt package.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
set retVal [catch {package require cawt} pkgVersion]

set appId [::Ppt::OpenNew]

puts [format "%-25s: %s" "Tcl version" [info patchlevel]]
puts [format "%-25s: %s" "Cawt version" $pkgVersion]
puts [format "%-25s: %s" "Twapi version" [::Cawt::GetPkgVersion "twapi"]]

puts [format "%-25s: %s (%s)" "PowerPoint Version" \
                             [::Ppt::GetVersion $appId] \
                             [::Ppt::GetVersion $appId true]]

puts [format "%-25s: %s" "PowerPoint extension" \
                             [::Ppt::GetExtString $appId]]

puts [format "%-25s: %s" "Active Printer" \
                        [::Cawt::GetActivePrinter $appId]]

puts [format "%-25s: %s" "User Name" \
                        [::Cawt::GetUserName $appId]]

puts [format "%-25s: %s" "Startup Pathname" \
                         [::Cawt::GetStartupPath $appId]]
puts [format "%-25s: %s" "Templates Pathname" \
                         [::Cawt::GetTemplatesPath $appId]]
puts [format "%-25s: %s" "Add-ins Pathname" \
                         [::Cawt::GetUserLibraryPath $appId]]
puts [format "%-25s: %s" "Installation Pathname" \
                         [::Cawt::GetInstallationPath $appId]]
puts [format "%-25s: %s" "User Folder Pathname" \
                         [::Cawt::GetUserPath $appId]]

set presId [::Ppt::AddPres $appId]

puts [format "%-30s: %s" "Appl. name (from Application)" \
         [::Cawt::GetApplicationName $appId]]
puts [format "%-30s: %s" "Appl. name (from Presentation)" \
         [::Cawt::GetApplicationName [::Cawt::GetApplicationId $presId]]]

::Ppt::Close $presId

if { [lindex $argv 0] eq "auto" } {
    ::Ppt::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
