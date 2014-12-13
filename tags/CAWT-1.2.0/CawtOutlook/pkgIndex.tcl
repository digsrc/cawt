# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc __CawtOutlookSourcePkgs { dir } {
    package provide cawtoutlook 1.0.0

    source [file join $dir outlookConst.tcl]
    source [file join $dir outlookBasic.tcl]
    rename ::__CawtOutlookSourcePkgs {}
}

# All modules are exported as package cawtoutlook
package ifneeded cawtoutlook 1.0.0 "[list __CawtOutlookSourcePkgs $dir]"
