# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { $argc == 0 } {
    set distDir [file join [pwd] ".." "CawtDistribution"]
} else {
    set distDir [lindex $argv 0]
}
set zipProg [auto_execok 7z.exe]
if { $zipProg eq "" } {
    set zipProg [file join "C:/" "opt" "7-Zip" "7z.exe"]
    if { ! [file exists $zipProg] } {
        set zipProg [file join "C:/" "Program Files" "7-Zip" "7z.exe"]
    }
}
if { $zipProg eq "" } {
    puts "No Zip program found. Exiting."
    exit 1
}

set cawtDir    [pwd]
set docDir     [file join $cawtDir "Documentation"]
set finalDir   [file join $docDir  "Final"]
set testDir    [file join $cawtDir "TestPrograms"]
set starkitDir [file join $cawtDir "Starkit"]

set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

set retVal [catch {package require cawt} cawtVersion]
if { $tcl_platform(pointerSize) == 4 } {
    set winPlatform "win32"
} else {
    set winPlatform "win64"
}

set distUserName "Cawt-$cawtVersion-User"
set distUserDir [file join $distDir $distUserName]

set distDevName "Cawt-$cawtVersion-Dev"
set distDevDir [file join $distDir $distDevName]

puts "Generating Cawt $cawtVersion distribution in $distDir"
file delete -force $distDir
file mkdir $distDir

puts "Cleaning test outputs and intermediate doc files ..."
cd $finalDir
file delete -force "CawtFigures"
cd $testDir
file delete -force "testOut"

cd $starkitDir
foreach winPlatform { "win32" "win64" } {
    set starkitName "CawtKit-$cawtVersion-$winPlatform.kit"
    puts "Copying starkit $starkitName to directory $distDir"
    file copy -force $starkitName $distDir
}

puts "Generating user distribution in directory $distUserDir"
file delete -force $distUserDir
file mkdir $distUserDir

cd $cawtDir
foreach dir { "CawtCore" "CawtEarth" "CawtExcel" "CawtExplorer" "CawtMatlab" \
              "CawtOcr" "CawtPpt" "CawtWord" "TestPrograms" } {
    puts "    Copying CAWT package $dir ..."
    file copy [file join $cawtDir $dir] $distUserDir
}

file mkdir [file join $distUserDir "Externals"]
foreach dir { "base64" "Img" "tablelist" "twapi"} {
    puts "    Copying external package $dir ..."
    file copy [file join $cawtDir "Externals" $dir] [file join $distUserDir "Externals"]
}
file copy pkgIndex.tcl $distUserDir

file copy $finalDir $distUserDir
cd $distUserDir
file rename "Final" "Documentation"

proc PackDir { dir zipName } {
    puts "    Generating $zipName from directory $dir ..."
    file delete -force $zipName

    exec $::zipProg a -tzip ${dir}.zip $dir -mx5
    file rename -force ${dir}.zip $zipName
}

cd $distDir
PackDir $distUserName "$distUserName.zip"

puts "Generating developer distribution in directory $distDevDir"
file delete -force $distDevDir
file copy [file join $cawtDir] $distDir
cd $distDir
file rename "Cawt" $distDevName

# Remove obsolete files and directories
cd $distDevName
file delete -force "_FOSSIL_"
file delete -force "Applications"
file delete -force "ToDo.txt"
cd "Starkit"
foreach winPlatform { "win32" "win64" } {
    set starkitName "CawtKit-$cawtVersion-$winPlatform.kit"
    file delete -force $starkitName
    file delete -force "CawtKit-$cawtVersion-$winPlatform.bat"
    file delete -force "CawtKit-$cawtVersion-$winPlatform.vfs"
}

cd $distDir
PackDir $distDevName "$distDevName.zip"

puts "Done"
