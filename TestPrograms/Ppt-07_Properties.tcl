# Test CawtPpt procedures related to property handling.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

# Open PowerPoint, show the application window and create a presentation.
set appId  [Ppt OpenNew]
set presId [Ppt AddPres $appId]

# Delete PowerPoint file from previous test run.
file mkdir testOut
set pptFile [file join [pwd] "testOut" "Ppt-07_Properties"]
append pptFile [Ppt GetExtString $appId]
file delete -force $pptFile

# Set some builtin and custom properties and check their values.
Cawt SetDocumentProperty $presId "Author"     "Paul Obermeier"
Cawt SetDocumentProperty $presId "Company"    "poSoft"
Cawt SetDocumentProperty $presId "Title"      $pptFile
Cawt SetDocumentProperty $presId "CustomProp" "CustomValue"

Cawt CheckString "Paul Obermeier" [Cawt GetDocumentProperty $presId "Author"]     "Property Author"
Cawt CheckString "poSoft"         [Cawt GetDocumentProperty $presId "Company"]    "Property Company"
Cawt CheckString $pptFile         [Cawt GetDocumentProperty $presId "Title"]      "Property Title"
Cawt CheckString "CustomValue"    [Cawt GetDocumentProperty $presId "CustomProp"] "Property CustomProp"

Cawt PrintNumComObjects

# Get all builtin and custom properties and insert them into the presentation.
set builtinSlide [Ppt AddSlide $presId]
set textboxId [Ppt AddTextbox $builtinSlide \
              [Cawt CentiMetersToPoints 1] [Cawt CentiMetersToPoints 2] \
              [Cawt CentiMetersToPoints 20] [Cawt CentiMetersToPoints 20]]
set builtinProps [Cawt GetDocumentProperties $presId "Builtin"]

Cawt PrintNumComObjects

foreach propertyName $builtinProps {
    Ppt AddTextboxText $textboxId "$propertyName: "
    Ppt AddTextboxText $textboxId [Cawt GetDocumentProperty $presId $propertyName] true
    incr row
}
Ppt SetTextboxFontSize $textboxId 10

Cawt PrintNumComObjects

set customSlide [Ppt AddSlide $presId]
set textboxId [Ppt AddTextbox $customSlide \
              [Cawt CentiMetersToPoints 1] [Cawt CentiMetersToPoints 2] \
              [Cawt CentiMetersToPoints 20] [Cawt CentiMetersToPoints 10]]
set customProps [Cawt GetDocumentProperties $presId "Custom"]

foreach propertyName [Cawt GetDocumentProperties $presId "Custom"] {
    Ppt AddTextboxText $textboxId "$propertyName: "
    Ppt AddTextboxText $textboxId [Cawt GetDocumentProperty $presId $propertyName] true
}
Ppt SetTextboxFontSize $textboxId 18

puts "Saving as PowerPoint file: $pptFile"
Ppt SaveAs $presId $pptFile

Cawt PrintNumComObjects

if { [lindex $argv 0] eq "auto" } {
    Ppt Quit $appId
    Cawt Destroy
    exit 0
}
Cawt Destroy
