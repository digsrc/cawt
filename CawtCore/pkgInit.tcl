# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc _InitCawtCore { dir version } {
    package provide cawtcore $version

    source [file join $dir cawtBasic.tcl]
    source [file join $dir cawtOffice.tcl]
    source [file join $dir cawtImgUtil.tcl]
    source [file join $dir cawtTestUtil.tcl]
}
