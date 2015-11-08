# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc __CawtMatlabSourcePkgs { dir } {
    package provide cawtmatlab 2.0.0

    source [file join $dir matlabBasic.tcl]
    rename ::__CawtMatlabSourcePkgs {}
}

# All modules are exported as package cawtmatlab
package ifneeded cawtmatlab 2.0.0 "[list __CawtMatlabSourcePkgs $dir]"