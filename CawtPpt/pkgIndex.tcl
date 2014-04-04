# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc __CawtPptSourcePkgs { dir } {
    package provide cawtppt 1.0.1

    source [file join $dir pptConst.tcl]
    source [file join $dir pptBasic.tcl]
    source [file join $dir pptUtil.tcl]
    rename ::__CawtPptSourcePkgs {}
}

# All modules are exported as package cawtppt
package ifneeded cawtppt 1.0.1 "[list __CawtPptSourcePkgs $dir]"
