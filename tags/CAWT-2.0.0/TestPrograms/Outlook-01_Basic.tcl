# Test basic functionality of the CawtOutlook package.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
set retVal [catch {package require cawt} pkgVersion]

set appId [Outlook OpenNew]

puts [format "%-25s: %s" "Tcl version" [info patchlevel]]
puts [format "%-25s: %s" "Cawt version" $pkgVersion]
puts [format "%-25s: %s" "Twapi version" [Cawt GetPkgVersion "twapi"]]
puts [format "%-25s: %s (%s)" "Outlook version" \
                             [Outlook GetVersion $appId] \
                             [Outlook GetVersion $appId true]]
puts ""
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

puts [format "%-30s: %s" "Appl. name (from Application)" [Cawt GetApplicationName $appId]]

puts [format "%-30s: %s" "Version (from Application)" [Outlook GetVersion $appId]]

puts ""
puts "Outlook has [llength [Outlook GetEnumTypes]] enumeration types."
set exampleEnum [lindex [Outlook GetEnumTypes] end]
puts "  $exampleEnum names : [Outlook GetEnumNames $exampleEnum]"
puts -nonewline "  $exampleEnum values:"
foreach name [Outlook GetEnumNames $exampleEnum] {
    puts -nonewline " [Outlook GetEnumVal $name]"
}

puts ""

if { [lindex $argv 0] eq "auto" } {
    Outlook Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
