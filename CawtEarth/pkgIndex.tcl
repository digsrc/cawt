# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc __CawtEarthSourcePkgs { dir } {
    package provide cawtearth 1.0.0

    source [file join $dir earthBasic.tcl]
    rename ::__CawtEarthSourcePkgs {}
}

# All modules are exported as package cawtge
package ifneeded cawtearth 1.0.0 "[list __CawtEarthSourcePkgs $dir]"
