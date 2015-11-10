# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

# Extend the auto_path to make Cawt subpackages available
if {[lsearch -exact $::auto_path $dir] == -1} {
    lappend ::auto_path $dir
}

proc _SetupCawtPkgs { dir version subDirs } {
    foreach subDir $subDirs {
        set pkg [string tolower $subDir]
        set subDirPath [file join $dir $subDir]
        set initFile [file join $subDirPath "pkgInit.tcl"]
        if { [file readable $initFile] } {
            source $initFile
            package ifneeded $pkg $version [list _Init$subDir $subDirPath $version]
        }
    }
}

proc _LoadCawtPkgs { dir version subDirs } {
    foreach subDir $subDirs {
        set pkg [string tolower $subDir]
        set retVal [catch { package require $pkg } ::__cawtPkgInfo($pkg,version)]
        set ::__cawtPkgInfo($pkg,avail) [expr !$retVal]
    }
    package provide cawt $version
}

set _CawtSubDirs [list CawtCore CawtExcel CawtWord CawtPpt CawtOutlook \
                       CawtOcr CawtExplorer CawtEarth CawtMatlab]

_SetupCawtPkgs $dir 2.1.2 $_CawtSubDirs

package ifneeded cawt 2.1.2 [list _LoadCawtPkgs $dir 2.1.2 $_CawtSubDirs]
