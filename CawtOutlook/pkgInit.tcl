# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc _InitCawtOutlook { dir version } {
    package provide cawtoutlook $version

    source [file join $dir outlookConst.tcl]
    source [file join $dir outlookBasic.tcl]
}
