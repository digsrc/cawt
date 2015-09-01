# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

set opts(Verbose)     false
set opts(RunTests)    true
set opts(RunCoverage) true

set nsUseList   [list]
set nsAvailList [list Excel Ppt Word Outlook Ocr Earth Matlab Explorer]
set useAllNamespaces false

proc PrintUsage { progName { msg "" } } {
    global opts nsAvailList

    puts ""
    if { $msg ne "" } {
        puts "Error: $msg"
    }
    puts ""
    puts "Usage: $progName \[Options\] Namespace \[Namespace\]"
    puts ""
    puts "Run the test programs and code coverage checks for specified namespace(s)."
    puts "Namespaces usable: $nsAvailList."
    puts "Use \"all\" as namespace name to run tests and checks for all namespaces."
    puts ""
    puts "Options:"
    puts "  --help   : Display this usage message and exit."
    puts "  --verbose: Show the detailed results of the tests. (Default: No)"
    puts "  --notests: Do not run the tests. (Default: Run tests)"
    puts "  --nocover: Do not run the coverage checks. (Default: Run coverage)"
    puts ""
}

proc RunTest { testFile } {
    global opts

    puts "Running test $testFile ..."
    set catchVal [catch {exec -ignorestderr $::tclExe $testFile auto 2>@1 } retVal optionsDict]
    if { $catchVal || [string match "*Error:*" $retVal] } {
        if { $catchVal } {
            set fullErrorInfo [dict get $optionsDict -errorinfo]
            set msgEndIndex [string first "\n" $fullErrorInfo]
            set msg [string range $fullErrorInfo 0 [expr {$msgEndIndex -1}]]
        } else {
            foreach line [split $retVal "\n"] {
                if { [string match "*Error:*" $line] } {
                    append msg $line
                }
            }
        }
        puts "Test $testFile failed: $msg"
    } else {
        if { $opts(Verbose) } {
            puts $retVal
            puts ""
        }
    }
}

set curArg 0
while { $curArg < $argc } {
    set curParam [lindex $argv $curArg]
    if { [string compare -length 1 $curParam "-"]  == 0 || \
         [string compare -length 2 $curParam "--"] == 0 } {
        set curOpt [string tolower [string trimleft $curParam "-"]]
        if { $curOpt eq "verbose" } {
            set opts(Verbose) true
        } elseif { $curOpt eq "help" } {
            PrintUsage $argv0
            exit 1
        } elseif { $curOpt eq "notests" } {
            set opts(RunTests) false
        } elseif { $curOpt eq "nocover" } {
            set opts(RunCoverage) false
        } else {
            PrintUsage $argv0 "Invalid option \"$curParam\" specified."
            exit 1
        }
    } else {
        if { $curParam eq "all" } {
            set useAllNamespaces true
        } else {
            lappend nsUseList $curParam
        }
    }
    incr curArg
}

if { $useAllNamespaces } {
    set nsUseList $nsAvailList
}

if { [llength $nsUseList] == 0 } {
    PrintUsage $argv0 "No namespace specified."
    exit 1
}

foreach nsName $nsUseList {
    if { [lsearch $nsAvailList $nsName] < 0 } {
        PrintUsage $argv0 "Invalid namespace \"$nsName\" specified."
        exit 1
    }
}

set tclExe [info nameofexecutable]

if { $opts(RunTests) } {
    catch { file mkdir testOut }
    foreach nsName $nsUseList {
        foreach f [lsort [glob ${nsName}-*]] {
            RunTest $f
        }
        puts ""
    }
}

if { $opts(RunCoverage) } {
    set cawtDir [file join [pwd] ".."]
    set auto_path [linsert $auto_path 0 $cawtDir [file join $cawtDir "Externals"]]

    package require cawt

    foreach nsName $nsUseList {
        puts "Checking $nsName test coverage ..."

        set procList [list]
        set allProcList [lsort [info commands ${nsName}::*]]
        foreach cmd $allProcList {
            if { ! [string match "*Obsolete:*" [info body $cmd]] } {
                if { [string first "::" $cmd] == 0 } {
                    set cmd [string range $cmd 2 end]
                }
                lappend procList $cmd
            }
        }

        # We search the test scripts as well as the implementation files,
        # as a procedure may be used by a higher-level procedure and thus
        # does not have to be tested separately.
        set nsLower [string tolower $nsName]
        set testFileList [lsort [glob "${nsName}-*.tcl" "../Cawt${nsName}/${nsLower}*.tcl"]]

        foreach testFile $testFileList {
            if { $opts(Verbose) } {
                puts "Scanning testfile $testFile"
            }
            set fp [open $testFile "r"]
            while { [gets $fp line] >= 0 } {
                foreach cmd $procList {
                    set ens [string map { "::" " " } $cmd]
                    if { [string match "*${cmd}*" $line] || [string match "*${ens}*" $line] } {
                        #puts "Found proc $cmd in file $testFile"
                        set found($cmd) 1
                    }
                }
            }
            close $fp
        }

        set foundList [lsort [array names found *]]
        foreach cmd $procList {
            if { [lsearch $foundList $cmd] < 0 } {
                puts "$cmd not yet tested"
            }
        }

        set numObsolete [expr [llength $allProcList] - [llength $procList]]
        puts "[llength $procList] procedures checked ($numObsolete obsolete)"
        puts ""
        unset found
    }
}

exit 0
