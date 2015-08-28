# Copyright: 2011-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc _InitCawtExplorer { dir version } {
    package provide cawtexplorer $version

    source [file join $dir explorerBasic.tcl]
}
