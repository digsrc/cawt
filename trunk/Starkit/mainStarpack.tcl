package require starkit

if { [starkit::startup] eq "sourced" } {
    return
}

package require cawt

if { $argc == 0 } {
    puts "Usage: $argv0 CAWT-Script"
    exit 0
} else {
    # If other command line parameters are supplied, assume the first
    # is the name of a Tcl script, which will be sourced.
    # Decrement the command argument counter and remove the sourced file
    # name from the command line parameter list.
    set i 0
    set tclScript [file normalize [lindex $argv $i]]
    incr argc -1
    set argv [lrange $argv [expr $i+1] end]
    set argv0 [file dirname [vfs::filesystem fullynormalize $argv0]]
    source $tclScript
}
