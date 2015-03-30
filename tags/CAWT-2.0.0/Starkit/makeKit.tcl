# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

# Call with all as parameter to build the CAWT StarKit for native OS.
# To select the StarKit version, specify "win32" or "win64" as 2nd parameter.
# Call with clean as parameter to remove all intermediate build files.
# Call with distclean as parameter to remove the StarKit, too.

set option "all"
if { $argc > 0 } {
    set option [lindex $argv 0]
}
set optClean false
if { $option eq "distclean" || $option eq "clean" } {
    set optClean true
}

set starkitDir [pwd]
set cawtDir    [file join $starkitDir ".."]

if { $tcl_platform(pointerSize) == 4 } {
    set inPlatform "win32"
} else {
    set inPlatform "win64"
}
set outPlatform $inPlatform
if { $argc > 1 } {
    set outOption [lindex $argv 1]
    if { $outOption eq "win32" || $outOption eq "win64" } {
        set outPlatform $outOption
    }
}
set tclkit "tclkit-sh-$inPlatform-86.exe"

set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

set retVal [catch {package require cawt} cawtVersion]

set starkitName      "CawtKit-$cawtVersion-$outPlatform"
set starkitVfs       [format "%s.vfs" $starkitName]
set starkitVfsDir    [file join $starkitDir $starkitVfs]
set starkitVfsLibDir [file join $starkitVfsDir "lib"]

file delete -force $starkitName.bat
file delete -force $starkitVfsDir

if { $option eq "distclean" } {
    file delete -force $starkitName.kit
}

if { $optClean } {
    puts "Cleaned $starkitVfsDir"
    if { $option eq "distclean" } {
        puts "Cleaned $starkitName.kit"
    }
    exit 0
}

puts "Generating StarKit $starkitName ..."

file mkdir $starkitVfsDir
file mkdir $starkitVfsLibDir

cd $cawtDir
foreach dir { "CawtCore" "CawtEarth" "CawtExcel" "CawtExplorer" "CawtMatlab" \
              "CawtOcr" "CawtOutlook" "CawtPpt" "CawtWord" } {
    puts "    Copying CAWT package $dir ..."
    file copy [file join $cawtDir $dir] [file join $starkitVfsLibDir $dir]
}
file copy pkgIndex.tcl $starkitVfsLibDir

foreach dir { "base64" "tablelist" "twapi" } {
    puts "    Copying external package $dir ..."
    file copy [file join $cawtDir "Externals" $dir] [file join $starkitVfsLibDir $dir]
}
# Handle Img package seperately.
set ImgPkg [file join "Img" "Img-$outPlatform"]
puts "    Copying external package $ImgPkg ..."
file copy [file join $cawtDir "Externals" $ImgPkg] [file join $starkitVfsLibDir "Img"]

cd $starkitDir
file copy main.tcl $starkitVfsDir

puts "Wrapping with $tclkit"
exec $tclkit sdx.kit wrap $starkitName.kit

puts "Done"
