# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

# Extend the auto_path to make Cawt subpackages available
if {[lsearch -exact $::auto_path $dir] == -1} {
    lappend ::auto_path $dir
}

proc __cawtSourcePkgs { dir } {
    set subPkgs [list cawtcore cawtexcel cawtword cawtppt cawtocr cawtexplorer \
                      cawtearth cawtmatlab]
    foreach pkg $subPkgs {
        set retVal [catch {package require $pkg} ::__cawtPkgInfo($pkg,version)]
        set ::__cawtPkgInfo($pkg,avail) [expr !$retVal]
    }
    package provide cawt 1.0.7
}

package ifneeded cawt 1.0.7 "[list __cawtSourcePkgs $dir]"
