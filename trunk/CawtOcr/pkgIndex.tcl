# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc __CawtOcrSourcePkgs { dir } {
    package provide cawtocr 1.0.0

    source [file join $dir ocrBasic.tcl]
    rename ::__CawtOcrSourcePkgs {}
}

# All modules are exported as package cawtocr
package ifneeded cawtocr 1.0.0 "[list __CawtOcrSourcePkgs $dir]"
