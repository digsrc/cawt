# Copyright: 2011-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc __CawtExplorerSourcePkgs { dir } {
    package provide cawtexplorer 1.0.0

    source [file join $dir explorerBasic.tcl]
    rename ::__CawtExplorerSourcePkgs {}
}

# All modules are exported as package cawtge
package ifneeded cawtexplorer 1.0.0 "[list __CawtExplorerSourcePkgs $dir]"
