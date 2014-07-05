# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc __CawtWordSourcePkgs { dir } {
    package provide cawtword 1.0.4

    source [file join $dir wordConst.tcl]
    source [file join $dir wordBasic.tcl]
    source [file join $dir wordUtil.tcl]
    rename ::__CawtWordSourcePkgs {}
}

# All modules are exported as package cawtword
package ifneeded cawtword 1.0.4 "[list __CawtWordSourcePkgs $dir]"
