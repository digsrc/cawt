# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.
#
# This script can be used to create all enumeration constants for the Cawt
# Office modules.
# See the createConstFile.bat batch file for using this script to generate
# the enumerations of Excel, Word and PowerPoint.

if { $argc != 2 } {
    puts ""
    puts "Usage: $argv0 FullPathOfOfficeApplication Namespace"
    puts ""
    exit 1
}

set applName [lindex $argv 0]
set nsName   [lindex $argv 1]

set shortApplName [join [lrange [file split $applName] 2 end] "/"]

set cawtDir [pwd]
set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

package require twapi 4

# Workaround for loading ITypeLibProxy_from_path in Twapi 4.0 alpha.
catch { twapi::name_to_iid }

set typeLib [twapi::ITypeLibProxy_from_path $applName]

set allEnumDict [dict get [$typeLib @Read -type enum] enum]

puts "# Auto generated by createConstFile.tcl based on the type library"
puts "# of \"$shortApplName\"."
puts "#"
puts "# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)"
puts "# Distributed under BSD license."
puts ""
puts "namespace eval $nsName \{"
puts ""
puts "    namespace ensemble create"

foreach enum [lsort [dict keys $allEnumDict]] {
    puts ""
    puts "    # Enumeration $enum"
    set enumDict [dict get $allEnumDict $enum]
    set valueDict [dict get $enumDict "-values"]
    foreach var [lsort [dict keys $valueDict]] {
        puts "    variable $var [dict get $valueDict $var]"
    }
}

puts ""
puts "    variable enums"
puts ""
puts "    array set enums \{"
foreach enum [lsort [dict keys $allEnumDict]] {
    puts -nonewline "        $enum \{"
    set enumDict [dict get $allEnumDict $enum]
    set valueDict [dict get $enumDict "-values"]
    foreach var [lsort [dict keys $valueDict]] {
        puts -nonewline " $var [dict get $valueDict $var]"
    }
    puts " \}"
}
puts "    \}"

set fp [open "constUtilProcs.tcl" "r"]
puts ""
puts [read $fp]
close $fp

puts "\}"

$typeLib Release

exit 0
