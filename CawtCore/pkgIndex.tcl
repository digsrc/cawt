# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc __CawtCoreSourcePkgs { dir } {
    package provide cawtcore 1.0.3

    source [file join $dir cawtBasic.tcl]
    source [file join $dir cawtImgUtil.tcl]
    rename ::__CawtCoreSourcePkgs {}
}

# All modules are exported as package cawtcore
package ifneeded cawtcore 1.0.3 "[list __CawtCoreSourcePkgs $dir]"
