# Test CawtExcel procedures to read data into a matrix and write matrix data
# into Matlab or RAW image files.
#
# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

set rawFile "testIn/gradient.raw"

set retVal [catch {package require img::raw} version]
if { $retVal == 0 } {
    set phImg1 [image create photo -file $rawFile \
                -format "RAW -useheader 1 -verbose 0 -nomap 0 -gamma 1"]
    pack [label .l1] -side left
    .l1 configure -image $phImg1
    wm title . "Original RAW image vs. generated RAW image \
               (Size: [image width $phImg1] x [image height $phImg1])"
    update
} else {
    puts "Package img::raw is not available."
    puts "Raw images will be read into a matrix, but not displayed in a Tk window."
}

# Delete files from previous test run.
file mkdir testOut
set rawOutFile [file join [pwd] "testOut" "Excel-10_Matrix.raw"]
file delete -force $rawOutFile
set matOutFile [file join [pwd] "testOut" "Excel-10_Matrix.mat"]
file delete -force $matOutFile

# Transfer RAW image file into a matrix and save as a Matlab file.
set t1 [clock clicks -milliseconds]
set matrixList1 [::Excel::ReadRawImageFile $rawFile]
::Excel::WriteMatlabFile $matrixList1 $matOutFile
set t2 [clock clicks -milliseconds]
puts "[expr $t2 - $t1] ms to put RAW image data into a Matlab file."

# Transfer Matlab file data into a matrix and save as a RAW image file.
set t1 [clock clicks -milliseconds]
set matrixList2 [::Excel::ReadMatlabFile $matOutFile]
::Excel::WriteRawImageFile $matrixList2 $rawOutFile
set t2 [clock clicks -milliseconds]
puts "[expr $t2 - $t1] ms to put Matlab matrix data into a RAW image file."
::Cawt::CheckMatrix $matrixList1 $matrixList2 "ReadMatlabFile"

if { $retVal == 0 } {
    set phImg2 [image create photo -file $rawOutFile \
                -format "RAW -useheader 1 -verbose 0 -nomap 0 -gamma 1"]
    pack [label .l2]
    .l2 configure -image $phImg2
}

::Cawt::Destroy
if { [lindex $argv 0] eq "auto" } {
    exit 0
}
