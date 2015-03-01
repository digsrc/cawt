# Test CawtMatlab procedures for executing Matlab commands.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

set mFile [file join [pwd] "testIn" "Expression.m"]
set fp [open $mFile]
set mContent [read $fp]
close $fp

set appId [Matlab OpenNew]

# Extract directory and pure file name.
set dirName  [file dirname $mFile]
set fileName [file tail $mFile]
set cmdName  [file rootname $fileName]

# Change working directory in Matlab.
Matlab ExecCmd $appId "cd $dirName"

# Load the specified Matlab file.
puts "Executing M-File $mFile"
puts $mContent
set result [Matlab ExecCmd $appId "$cmdName"]
puts $result

if { [lindex $argv 0] eq "auto" } {
    Matlab Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
