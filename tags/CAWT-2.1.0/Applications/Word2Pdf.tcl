#!/bin/sh
# the next line restarts using tclsh \
exec tclsh "$0" -- ${1+"$@"}

# Utility script to convert a Word document to a PDF file.
#
# Copyright: 2013-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

set cawtDir [file join [pwd] ".."]
set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

package require cawt

proc PrintUsage { appName } {
    puts ""
    puts "Usage: $appName \[Options\] WordFile PdfFile"
    puts ""
    puts "Options: None at the moment."
    puts ""
}

set optPrintHelp false
set wordFile     ""
set pdfFile      ""

set curArg 0
while { $curArg < $argc } {
    set curParam [lindex $argv $curArg]
    if { [string compare -length 1 $curParam "-"]  == 0 || \
         [string compare -length 2 $curParam "--"] == 0 } {
        set curOpt [string tolower [string trimleft $curParam "-"]]
        if { $curOpt eq "help" } {
            set optPrintHelp true
        }
    } else {
        if { $wordFile eq "" } {
            set wordFile $curParam
        } elseif { $pdfFile eq "" } {
            set pdfFile $curParam
        }
    }
    incr curArg
}

if { $optPrintHelp } {
    PrintUsage $argv0
    exit 0
}

# Check, if all necessary parameters have been supplied.
if { $wordFile eq "" } {
    puts "No Word input file specified."
    PrintUsage $argv0
    exit 1
}
if { $pdfFile eq "" } {
    puts "No PDF output file specified."
    PrintUsage $argv0
    exit 1
}

set wordFile [file nativename [file normalize $wordFile]]
set pdfFile  [file nativename [file normalize $pdfFile]]

if { ! [file exists $wordFile] } {
    puts "Specified Word file $wordFile not existent."
    PrintUsage $argv0
    exit 1
}

puts "Reading Word file: $wordFile"
# Open new Word instance and show the application window.
set appId [Word OpenNew true]

# Delete PDF file, if existent.
# file delete -force $pdfFile

# Open the Word document in read-only mode.
set docId [Word OpenDocument $appId $wordFile true]

puts "Saving as PDF file: $pdfFile"
# # Use in a catch statement, as PDF export is available only in Word 2007 and up.
set catchVal [ catch { Word SaveAsPdf $docId $pdfFile } retVal]
if { $catchVal } {
    puts "Error: $retVal"
}

# Quit Word application without showing possible alerts.
Word Close $docId
Word Quit $appId false
Cawt Destroy
exit 0
