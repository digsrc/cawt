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
set tclkit     "tclkit-$inPlatform-tcl.exe"
set runtimeTcl "tclkit-$inPlatform-sh.exe"

set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

set retVal [catch {package require cawt} cawtVersion]

set starpackName     "Cawt-$cawtVersion-$outPlatform"
set starkitName      "CawtKit-$cawtVersion-$outPlatform"
set starkitVfs       [format "%s.vfs" $starkitName]
set starkitVfsDir    [file join $starkitDir $starkitVfs]
set starkitVfsLibDir [file join $starkitVfsDir "lib"]

file delete -force $starkitName.bat
file delete -force $starkitVfsDir
file delete -force $runtimeTcl

if { $option eq "distclean" } {
    file delete -force $starkitName.kit
    file delete -force $starpackName.exe
}

if { $optClean } {
    puts "Cleaned $starkitVfsDir"
    if { $option eq "distclean" } {
        puts "Cleaned $starkitName.kit"
        puts "Cleaned $starpackName.exe"
    }
    exit 0
}

file copy $tclkit $runtimeTcl

puts "Generating StarKit $starkitName.kit ..."

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
file copy mainStarkit.tcl [file join $starkitVfsDir "main.tcl"]

puts "    Wrapping with $tclkit ..."
exec $tclkit sdx.kit wrap $starkitName.kit

puts "Generating Starpack $starkitName.exe"
file delete -force $starpackName.exe
file delete -force [file join $starkitVfsDir "main.tcl"]

set batchDir "${outPlatform}Batch"
puts "    Copying Tk package $batchDir ..."
foreach dir [glob $batchDir/*] {
    puts "        Copying $dir ... "
    file copy $dir [file join $starkitVfsLibDir]
}

file copy mainStarpack.tcl [file join $starkitVfsDir "main.tcl"]

puts "    Using runtime $runtimeTcl"
exec $tclkit sdx.kit wrap $starkitName.exe -runtime $runtimeTcl
file rename $starkitName.exe $starpackName.exe

file delete -force $runtimeTcl
puts "Done"
