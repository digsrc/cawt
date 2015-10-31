# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc _InitCawtPpt { dir version } {
    package provide cawtppt $version

    source [file join $dir pptConst.tcl]
    source [file join $dir pptBasic.tcl]
    source [file join $dir pptUtil.tcl]
}
