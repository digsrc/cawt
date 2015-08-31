package require starkit

if { [starkit::startup] eq "sourced" } {
    return
}

if { $tcl_platform(platform) eq "windows" } {
    package require Tk
    console show
}

package require cawt
