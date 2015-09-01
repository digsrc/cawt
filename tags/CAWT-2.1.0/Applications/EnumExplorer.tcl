#!/bin/sh
# the next line restarts using tclsh \
exec tclsh "$0" -- ${1+"$@"}

# Utility script to convert a Word document to a PDF file.
#
# Copyright: 2013-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

set cawtDir [file join [pwd] ".."]
set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

package require Tk
package require tablelist
package require cawt

set gPo(defNamespaces) [list "Excel" "Ppt" "Word" "Outlook"]
set startNamespace [lindex $gPo(defNamespaces) 0]

proc PrintUsage { appName } {
    global gPo

    puts ""
    puts "Usage: $appName \[Namespace\]"
    puts ""
    puts "Start enumeration explorer and display specified namespace."
    puts "If no namespace is given, the first namespace in this list is used:"
    puts "  $gPo(defNamespaces)"
    puts ""
    puts "Pressing the \"Copy as enum\" button or the \"C\" key copies the"
    puts "currently selected enumeration as Tcl variable to the clipboard."
    puts "Example: \$Excel::xlAboveStdDev"
    puts ""
    puts "Pressing the \"Copy as string\" button or the \"c\" key copies the"
    puts "currently selected enumeration name to the clipboard."
    puts "Example: xlAboveStdDev"
    puts ""
}

proc MouseWheelCB { w delta dir } {
    if {[tk windowingsystem] ne "aqua"} {
        set delta [expr {($delta / 120) * 4}]
    }
    if { $dir eq "y" } {
        $w yview scroll [expr {-$delta}] units
    } else {
        $w xview scroll [expr {-$delta}] units
    }
}

proc CreateScrolledWidget { wType w titleStr args } {
    if { [winfo exists $w.par] } {
        destroy $w.par
    }
    ttk::frame $w.par
    pack $w.par -side top -fill both -expand 1
    if { $titleStr ne "" } {
        label $w.par.label -text "$titleStr" -anchor center
    }
    $wType $w.par.widget \
           -xscrollcommand "$w.par.xscroll set" \
           -yscrollcommand "$w.par.yscroll set" {*}$args
    ttk::scrollbar $w.par.xscroll -command "$w.par.widget xview" -orient horizontal
    ttk::scrollbar $w.par.yscroll -command "$w.par.widget yview" -orient vertical
    set rowNo 0
    if { $titleStr ne "" } {
        set rowNo 1
        grid $w.par.label -sticky ew -columnspan 2
    }
    grid $w.par.widget $w.par.yscroll -sticky news
    grid $w.par.xscroll               -sticky ew

    grid rowconfigure    $w.par $rowNo -weight 1
    grid columnconfigure $w.par 0      -weight 1

    bind $w.par.widget <MouseWheel>       "MouseWheelCB $w.par.widget %D y"
    bind $w.par.widget <Shift-MouseWheel> "MouseWheelCB $w.par.widget %D x"

    return $w.par.widget
}

proc SetScrolledTitle { w titleStr } {
    set pathList [split $w "."]
    # Index -3 is needed for CreateScrolledFrame.
    # Index -2 is needed for all other widget types.
    foreach ind { -2 -3 } {
        set parList  [lrange $pathList 0 [expr [llength $pathList] $ind]]
        set parPath  [join $parList "."]

        set labelPath $parPath
        append labelPath ".label"
        if { [winfo exists $labelPath] } {
            $labelPath configure -text $titleStr
            break
        }
    }
}

proc CreateScrolledTablelist { w titleStr args } {
    return [CreateScrolledWidget tablelist::tablelist $w $titleStr {*}$args]
}

proc WriteInfoStr { msg } {
    $::statFr.l configure -text "$msg"
}

proc ResetEnumTable {} {
    global gPo

    $gPo(enumTable) delete 0 end
    SetScrolledTitle $gPo(enumTable) "Enumerations"
}

proc ResetValTable {} {
    global gPo

    $gPo(valTable) delete 0 end
    set gPo(curEnumName) ""
    SetScrolledTitle $gPo(valTable) "Values"
}

proc ShowEnums { ns } {
    global gPo

    set enumTypes [${ns}::GetEnumTypes]
    ResetEnumTable
    ResetValTable
    foreach enum $enumTypes {
        $gPo(enumTable) insert end [list "" $enum]
    }
    SetScrolledTitle $gPo(enumTable) "$ns ([llength $enumTypes] enumerations)"
}

proc ShowEnumValues { w } {
    global gPo

    set rowList [$w curselection]
    if { [llength $rowList] > 0 } {
        ResetValTable
        set row [lindex $rowList 0]
        set enum [$w cellcget $row,1 -text]
        set fullEnum "$gPo(curNs)::$enum"
        set enumNames [$gPo(curNs)::GetEnumNames $enum]
        foreach enumName $enumNames {
            set val [$gPo(curNs)::GetEnumVal $enumName]
            $gPo(valTable) insert end [list $enumName $val]
        }
        SetScrolledTitle $gPo(valTable) "$fullEnum ([llength $enumNames] values)"
    }
}

proc SelectEnumValue { w } {
    global gPo

    set rowList [$w curselection]
    if { [llength $rowList] > 0 } {
        set row [lindex $rowList 0]
        set gPo(curEnumName) [$w cellcget $row,0 -text]
    }
}

proc CopyToClipboard { type } {
    global gPo

    if { $type eq "name" } {
        set val "$gPo(curEnumName) "
    } else {
        set val "\$$gPo(curNs)::$gPo(curEnumName) "
    }

    twapi::open_clipboard
    twapi::empty_clipboard
    twapi::write_clipboard 1 $val
    twapi::close_clipboard
    WriteInfoStr "Copied to clipboard: $val"
}

set optPrintHelp  false

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
        if { [lsearch -exact $gPo(defNamespaces) $curParam] >= 0 } {
            set startNamespace $curParam
        } else {
            puts "Warning: $curParam is not supported."
        }
    }
    incr curArg
}

if { $optPrintHelp } {
    PrintUsage $argv0
    exit 0
}

set gPo(curNs) $startNamespace
set gPo(curEnumName) ""

set tw .enumExplorer

toplevel $tw
wm withdraw .
wm title $tw "CAWT Enumeration Explorer"

set nsFr   $tw.ns
set enumFr $tw.enum
set valFr  $tw.val
set statFr $tw.stat

labelframe $nsFr
frame $enumFr
frame $valFr
frame $statFr -relief sunken -borderwidth 1

grid $nsFr   -row 0 -column 0 -sticky news -columnspan 2
grid $enumFr -row 1 -column 0 -sticky news
grid $valFr  -row 1 -column 1 -sticky news
grid $statFr -row 2 -column 0 -sticky news -columnspan 2
grid rowconfigure    $tw 1 -weight 1
grid columnconfigure $tw 1 -weight 1

ttk::label $nsFr.l -text "Show namespace:"
pack $nsFr.l -side left

foreach ns $gPo(defNamespaces) {
    ttk::radiobutton $nsFr.b$ns -style Toolbutton -text $ns \
                     -variable gPo(curNs) -value $ns -command "ShowEnums $ns"
    pack $nsFr.b$ns -side left -padx 3
}

ttk::button $nsFr.exit -style Toolbutton -text "Quit" -command "exit"
pack $nsFr.exit -side right -padx 10

ttk::button $nsFr.copy1 -style Toolbutton -text "Copy as string" -command "CopyToClipboard name"
ttk::button $nsFr.copy2 -style Toolbutton -text "Copy as enum"   -command "CopyToClipboard enum"
pack $nsFr.copy1 $nsFr.copy2 -side right -padx 3

bind $tw <Key-c>      "CopyToClipboard name"
bind $tw <Key-C>      "CopyToClipboard enum"
bind $tw <Key-Escape> exit

set gPo(enumTable) [CreateScrolledTablelist $enumFr "Enumerations" \
    -width 50 -exportselection false \
    -columns { 0 "#"           "right"
               0 "Enumeration" "left" } \
    -stretch 1 \
    -stripebackground #e0e8f0 \
    -labelcommand ::tablelist::sortByColumn \
    -showseparators yes]
$gPo(enumTable) columnconfigure 0 -showlinenumbers true
$gPo(enumTable) columnconfigure 1 -sortmode dictionary
bind $gPo(enumTable) <<ListboxSelect>> "ShowEnumValues %W"

set gPo(valTable) [CreateScrolledTablelist $valFr "Enumeration values" \
    -width 40 -exportselection false \
    -columns { 0 "Name"  "left"
               0 "Value" "right" } \
    -stretch 0 \
    -stripebackground #e0e8f0 \
    -labelcommand ::tablelist::sortByColumn \
    -showseparators yes]
$gPo(valTable) columnconfigure 0 -sortmode dictionary
$gPo(valTable) columnconfigure 1 -sortmode integer
bind $gPo(valTable) <<ListboxSelect>> "SelectEnumValue %W"

ttk::label $statFr.l -text ""
pack $statFr.l -side left -expand true -fill x

ShowEnums $gPo(curNs)
