# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc __CawtOcrSourcePkgs { dir } {
    package provide cawtocr 2.0.0

    source [file join $dir ocrBasic.tcl]
    rename ::__CawtOcrSourcePkgs {}
}

# All modules are exported as package cawtocr
package ifneeded cawtocr 2.0.0 "[list __CawtOcrSourcePkgs $dir]"
