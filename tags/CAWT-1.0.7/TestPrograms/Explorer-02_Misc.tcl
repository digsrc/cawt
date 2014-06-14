# Test miscellaneous CawtExplorer functions like navigating to an URL and using fillscreen mode.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

set appId [::Explorer::OpenNew]

set htmlFile [file join [pwd] "testIn" "WikitTable.txt"]
set url "http://www.posoft.de/html/extTcomExcel.html"

::Explorer::Navigate $appId $htmlFile false
puts "IsBusy: [::Explorer::IsBusy $appId]"
after 500
::Explorer::Navigate $appId $url
puts "IsBusy: [::Explorer::IsBusy $appId]"

::Explorer::Go $appId "Back"
::Explorer::FullScreen $appId true
after 1000
::Explorer::Go $appId "Forward"
after 1000
::Explorer::FullScreen $appId false

if { [lindex $argv 0] eq "auto" } {
    ::Explorer::Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
