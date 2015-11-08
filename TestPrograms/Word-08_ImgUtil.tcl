# Test CawtWord procedures for dealing with images.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

if { [file exists "SetTestPathes.tcl"] } {
    source "SetTestPathes.tcl"
}
package require cawt

set squareImg    [file join [pwd] "testIn/Square.gif"]
set landscapeImg [file join [pwd] "testIn/Landscape.gif"]
set portraitImg  [file join [pwd] "testIn/Portrait.gif"]
set wishImg      [file join [pwd] "testIn/wish.gif"]

# Open new Word instance and show the application window.
set appId [Word OpenNew true]

# Delete Word file from previous test run.
file mkdir testOut
set wordFile [file join [pwd] "testOut" "Word-08_ImgUtil"]
append wordFile [Word GetExtString $appId]
file delete -force $wordFile

# Create a new document.
set docId [Word AddDocument $appId]

puts "Inserting images of different sizes ..."
Word AppendText $docId "Images of different sizes\n"

Word AppendText $docId "Square image:\n"
Word InsertImage [Word GetEndRange $docId] $squareImg
Word AppendParagraph $docId

Word AppendText $docId "Landscape image:\n"
Word InsertImage [Word GetEndRange $docId] $landscapeImg
Word AppendParagraph $docId

Word AppendText $docId "Portrait image:\n"
Word InsertImage [Word GetEndRange $docId] $portraitImg
Word AppendParagraph $docId

puts "Inserting images with different modes ..."
Word AddPageBreak [Word GetEndRange $docId]
Word AppendText $docId "Images with different insertion modes\n"

Word AppendText $docId "Linked image:\n"
Word InsertImage [Word GetEndRange $docId] $squareImg true false
Word AppendParagraph $docId

Word AppendText $docId "Embedded image:\n"
Word InsertImage [Word GetEndRange $docId] $squareImg false true
Word AppendParagraph $docId

Word AppendText $docId "Linked and Embedded image:\n"
Word InsertImage [Word GetEndRange $docId] $squareImg false true
Word AppendParagraph $docId

set catchVal [ catch { Word InsertImage [Word GetEndRange $docId] $squareImg false false } retVal]
if { $catchVal } {
    puts "Successfully caught: $retVal"
}

puts "Inserting and scaling images ..."
Word AddPageBreak [Word GetEndRange $docId]
Word AppendText $docId "Images with different scalings\n"

Word AppendText $docId "Landscape scaled to Square:\n"
set scaleId1 [Word InsertImage [Word GetEndRange $docId] $landscapeImg]
Word ScaleImage $scaleId1 1 2
Word AppendParagraph $docId

Word AppendText $docId "Portrait scaled to Square:\n"
set scaleId2 [Word InsertImage [Word GetEndRange $docId] $portraitImg]
Word ScaleImage $scaleId2 2 1
Word AppendParagraph $docId

puts "Inserting and cropping images ..."
Word AddPageBreak [Word GetEndRange $docId]
Word AppendText $docId "Images with different croppings\n"
# CropImage shapeId cropBottom cropTop cropLeft cropRight

Word AppendText $docId "Square cropped at the bottom side:\n"
set cropId1 [Word InsertImage [Word GetEndRange $docId] $squareImg]
Word CropImage $cropId1 5c 0  0 0
Word AppendParagraph $docId

Word AppendText $docId "Square cropped at the top side:\n"
set cropId2 [Word InsertImage [Word GetEndRange $docId] $squareImg]
Word CropImage $cropId2 0 0.5c  0 0
Word AppendParagraph $docId

Word AppendText $docId "Square cropped at the left side:\n"
set cropId3 [Word InsertImage [Word GetEndRange $docId] $squareImg]
Word CropImage $cropId3 0 0  2c 0
Word AppendParagraph $docId

Word AppendText $docId "Square cropped at the right side:\n"
set cropId4 [Word InsertImage [Word GetEndRange $docId] $squareImg]
Word CropImage $cropId4 0 0  0 2c
Word AppendParagraph $docId

# Save document as Word file.
puts "Saving as Word file: $wordFile"
Word SaveAs $docId $wordFile

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Word Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
