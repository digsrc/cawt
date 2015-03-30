# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc __CawtCoreSourcePkgs { dir } {
    package provide cawtcore 2.0.0

    source [file join $dir cawtBasic.tcl]
    source [file join $dir cawtOffice.tcl]
    source [file join $dir cawtImgUtil.tcl]
    source [file join $dir cawtTestUtil.tcl]
    rename ::__CawtCoreSourcePkgs {}
}

# All modules are exported as package cawtcore
package ifneeded cawtcore 2.0.0 "[list __CawtCoreSourcePkgs $dir]"
