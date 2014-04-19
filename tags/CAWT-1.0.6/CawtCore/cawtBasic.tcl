# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval ::Cawt {
    variable pkgInfo
    variable dotsPerInch
    variable _ruffdoc

    lappend _ruffdoc Introduction {
        The Cawt namespace provides commands for basic automation functionality.
    }

    proc _Init {} {
        variable pkgInfo
        variable dotsPerInch

        set dotsPerInch 72

        set retVal [catch {package require twapi 4} version]
        set pkgInfo(twapi,avail)   [expr !$retVal]
        set pkgInfo(twapi,version) $version

        set retVal [catch {package require cawtcore} version]
        set pkgInfo(cawtcore,avail)   [expr !$retVal]
        set pkgInfo(cawtcore,version) $version

        set retVal [catch {package require cawtearth} version]
        set pkgInfo(cawtearth,avail)   [expr !$retVal]
        set pkgInfo(cawtearth,version) $version

        set retVal [catch {package require cawtexcel} version]
        set pkgInfo(cawtexcel,avail)   [expr !$retVal]
        set pkgInfo(cawtexcel,version) $version

        set retVal [catch {package require cawtexplorer} version]
        set pkgInfo(cawtexplorer,avail)   [expr !$retVal]
        set pkgInfo(cawtexplorer,version) $version

        set retVal [catch {package require cawtmatlab} version]
        set pkgInfo(cawtmatlab,avail)   [expr !$retVal]
        set pkgInfo(cawtmatlab,version) $version

        set retVal [catch {package require cawtocr} version]
        set pkgInfo(cawtocr,avail)   [expr !$retVal]
        set pkgInfo(cawtocr,version) $version

        set retVal [catch {package require cawtppt} version]
        set pkgInfo(cawtppt,avail)   [expr !$retVal]
        set pkgInfo(cawtppt,version) $version

        set retVal [catch {package require cawtword} version]
        set pkgInfo(cawtword,avail)   [expr !$retVal]
        set pkgInfo(cawtword,version) $version
    }

    proc HavePkg { pkgName } {
        # Check, if a Cawt sub-package is available.
        #
        # pkgName - The name of the sub-package.
        #
        # Return true, if sub-package pkgName was loaded successfully.
        # Otherwise return false.
        #
        # See also: GetPkgVersion

        variable pkgInfo

        if { [info exists pkgInfo($pkgName,avail)] } {
            return $pkgInfo($pkgName,avail)
        }
        return false
    }

    proc GetPkgVersion { pkgName } {
        # Get the version of a Cawt sub-package.
        #
        # pkgName - The name of the sub-package
        #
        # The version is returned as a string.
        #
        # See also: HavePkg

        variable pkgInfo

        set retVal ""
        if { [HavePkg $pkgName] } {
            set retVal $pkgInfo($pkgName,version)
        }
        return $retVal
    }

    proc SetDotsPerInch { dpi } {
        # Set the dots-per-inch value used for conversions.
        #
        # dpi - Integer dpi value.
        #
        # If the dpi value is not explicitely set with this function, it's
        # default value is 72.
        #
        # No return value.
        #
        # See also: GetDotsPerInch

        variable dotsPerInch

        set dotsPerInch $dpi
    }

    proc GetDotsPerInch {} {
        # Return the dots-per-inch value used for conversions.
        #
        # See also: SetDotsPerInch

        variable dotsPerInch

        return $dotsPerInch
    }

    proc InchesToPoints { inches } {
        # Convert inch value into points.
        #
        # inches - Floating point inch value to be converted to points.
        #
        # Return the corresponding value in points.
        #
        # See also: SetDotsPerInch CentiMetersToPoints

        variable dotsPerInch

        return [expr {$inches * double($dotsPerInch)}]
    }

    proc CentiMetersToPoints { cm } {
        # Convert centimeter value into points.
        #
        # cm - Floating point centimeter value to be converted to points.
        #
        # Return the corresponding value in points.
        #
        # See also: SetDotsPerInch InchesToPoints

        variable dotsPerInch

        return [expr {$cm / 2.54 * double($dotsPerInch)}]
    }

    proc RgbToColor { r g b } {
        # Return a RGB color as an Office color number.
        #
        # r - The red component of the color
        # g - The green component of the color
        # b - The blue component of the color
        #
        # The r, g and b values are specified as integers in the
        # range 0 .. 255.
        #
        # Return the color number as an integer.

        return [expr {int ($b) << 16 | int ($g) << 8 | int($r)}]
    }

    proc TclInt { val } {
        # Cast a value to an integer with boolean range.
        #
        # val - The value to be casted.
        #
        # Return 1, if val is not equal to zero or true.
        # Return 0, if val is equal to zero or false.
        #
        # See also: TclBool

        set tmp 0
        if { $val } {
            set tmp 1
        }
        return $tmp
    }

    proc TclBool { val } {
        # Cast a value to a boolean.
        #
        # val - The value to be casted.
        #
        # Return true, if val is not equal to zero or true.
        # Return false, if val is equal to zero or false.
        #
        # See also: TclInt

        return [twapi::tclcast boolean $val]
    }

    proc GetOrCreateApp { appName useExistingFirst } {
        # Use or create an instance of an application.
        #
        # appName          - The name of the application to be create or used.
        # useExistingFirst - Prefer an already running application.
        #
        # Application names supported and tested with Cawt are:
        # "Excel.Application", "PowerPoint.Application", "Word.Application",
        # "GoogleEarth.ApplicationGE", "Matlab.Application".
        # Note: There are higher level functions "Open" and "OpenNew" for the
        # Cawt sub-packages.
        #
        # If "useExistingFirst" is set to true, it is checked, if an application
        # instance is already running. If true, this instance is used.
        # If no running application is available, a new instance is started.
        #
        # Return the application identifier.
        #
        # See also: KillApp

        set foundApp false
        if { ! [HavePkg "twapi"] } {
            error "Cannot use $appName. No Twapi extension available."
        }
        if { $useExistingFirst } {
            set retVal [catch {twapi::comobj $appName -active} appId]
            if { $retVal == 0 } {
                set foundApp true
            }
        }
        if { $foundApp == false } {
            set retVal [catch {twapi::comobj $appName} appId]
        }
        if { $foundApp == true || $retVal == 0 } {
            return $appId
        }
        error "Cannot get or create $appName object."
    }

    proc KillApp { progName } {
        # Kill all running instances of an application.
        #
        # progName - The application's program name, as shown in the task manager.
        #
        # No return value.
        #
        # See also: GetOrCreateApp

        set pids [concat [twapi::get_process_ids -name $progName] \
                         [twapi::get_process_ids -path $progName]]
        foreach pid $pids {
            # Catch the error in case process does not exist any more
            catch {twapi::end_process $pid -force}
        }
    }

    proc ShowAlerts { appId onOff } {
        # Toggle the display of Office alerts.
        #
        # appId - The application identifier.
        # onOff - Switch the alerts on or off.
        #
        # No return value.

        if { $onOff } {
            if { [::Cawt::GetApplicationName $appId] eq "Microsoft Word" } {
                set alertLevel [expr $::Word::wdAlertsAll]
            } else {
                set alertLevel [expr 1]
            }
        } else {
            set alertLevel [expr 0]
        }
        $appId DisplayAlerts $alertLevel
    }

    proc IsValidId { comObj } {
        # Check, if a COM object is valid.
        #
        # comObj - The COM object.
        #
        # Return true, if "comObj" is a valid object.
        # Otherwise return false.

        return [expr { $comObj ne "" && ![$comObj -isnull] } ]
    }

    proc Destroy { { comObj "" } } {
        # Destroy one or all COM objects.
        #
        # comObj - The COM object to be destroyed.
        #
        # If "comObj" is an empty string, all existing COM objects are destroyed.
        # Otherwise only the specified COM object is destroyed.
        #
        # Note: Twapi does not clean up generated COM object identifiers, so you
        # have to put a call to Destroy at the end of your Cawt script.
        # For further details about COM objects and their lifetime see the Twapi
        # documentation.

        if { $comObj ne "" } {
            $comObj -destroy
        } else {
            # puts "Before: [llength [twapi::comobj_instances]]"
            foreach obj [twapi::comobj_instances] {
                $obj -destroy
            }
            # puts "After: [llength [twapi::comobj_instances]]"
        }
    }

    proc GetApplicationId { componentId } {
        # Get the application identifier of an Office component.
        #
        # componentId - The identifier of an Office component.
        #
        # Office components are Workbooks, Worksheets, ...

        return [$componentId Application]
    }

    proc GetApplicationName { appId } {
        # Get the name of an Office application.
        #
        # appId - The application identifier.
        #
        # Return the name of the application as a string.

        return [$appId Name]
    }

    proc GetActivePrinter { appId } {
        # Get the name of the active printer.
        #
        # appId - The application identifier.
        #
        # Return the name of the active printer as a string.

        set retVal [catch {$appId ActivePrinter} val]
        if { $retVal == 0 } {
            return $val
        } else {
            return "Method not available"
        }
    }

    proc GetUserName { appId } {
        # Get the name of the Office application user.
        #
        # appId - The application identifier.
        #
        # Return the name of the application user as a string.

        set retVal [catch {$appId UserName} val]
        if { $retVal == 0 } {
            return $val
        } else {
            return "Method not available"
        }
    }

    proc GetStartupPath { appId } {
        # Get the Office startup pathname.
        #
        # appId - The application identifier.
        #
        # Return the startup pathname as a string.

        set retVal [catch {$appId StartupPath} val]
        if { $retVal == 0 } {
            return $val
        } else {
            return "Method not available"
        }
    }

    proc GetTemplatesPath { appId } {
        # Get the Office templates pathname.
        #
        # appId - The application identifier.
        #
        # Return the templates pathname as a string.

        set retVal [catch {$appId TemplatesPath} val]
        if { $retVal == 0 } {
            return $val
        } else {
            return "Method not available"
        }
    }

    proc GetUserLibraryPath { appId } {
        # Get the Office user library pathname.
        #
        # appId - The application identifier.
        #
        # Return the user library pathname as a string.

        set retVal [catch {$appId UserLibraryPath} val]
        if { $retVal == 0 } {
            return $val
        } else {
            return "Method not available"
        }
    }

    proc GetInstallationPath { appId } {
        # Get the Office installation pathname.
        #
        # appId - The application identifier.
        #
        # Return the installation pathname as a string.

        set retVal [catch {$appId Path} val]
        if { $retVal == 0 } {
            return $val
        } else {
            return "Method not available"
        }
    }

    proc GetUserPath { appId } {
        # Get the Office user folder's pathname.
        #
        # appId - The application identifier.
        #
        # Return the user folder's pathname as a string.

        set retVal [catch {$appId DefaultFilePath} val]
        if { $retVal == 0 } {
            return $val
        } else {
            return "Method not available"
        }
    }

    proc CheckString { expected value msg { printCheck true } } {
        # Check, if 2 string values are identical.
        #
        # expected   - Expected string value.
        # value      - Test string value.
        # msg        - Message for test case.
        # printCheck - Print message for successful test case.
        #
        # Return true, if both string values are identical.
        # If "printCheck" is set to true, a line prepended with "Check:" and the
        # message supplied in "msg" is printed to standard output.
        # If the check fails, return false and print message prepended with "Error:".

        if { $expected ne $value } {
            puts "Error: $msg (Expected: \"$expected\" Have: \"$value\")"
            return false
        }
        if { $printCheck } {
            puts "Check: $msg (Expected: \"$expected\" Have: \"$value\")"
        }
        return true
    }

    proc CheckNumber { expected value msg { printCheck true } } {
        # Check, if 2 numerical values are identical.
        #
        # expected   - Expected numeric value.
        # value      - Test numeric value.
        # msg        - Message for test case.
        # printCheck - Print message for successful test case.
        #
        # Return true, if both numeric values are identical.
        # If "printCheck" is set to true, a line prepended with "Check:" and the
        # message supplied in "msg" is printed to standard output.
        # If the check fails, return false and print message prepended with "Error:".

        if { $expected != $value } {
            puts "Error: $msg (Expected: $expected Have: $value)"
            return false
        }
        if { $printCheck } {
            puts "Check: $msg (Expected: $expected Have: $value)"
        }
        return true
    }

    proc CheckList { expected value msg { printCheck true } } {
        # Check, if 2 lists are identical.
        #
        # expected   - Expected list.
        # value      - Test list.
        # msg        - Message for test case.
        # printCheck - Print message for successful test case.
        #
        # Return true, if both lists are identical.
        # If "printCheck" is set to true, a line prepended with "Check:" and the
        # message supplied in "msg" is printed to standard output.
        # If the check fails, return false and print message prepended with "Error:".

        if { [llength $expected] != [llength $value] } {
            puts "Error: $msg (List length differ. Expected: [llength $expected] Have: [llength $value])"
            return false
        }
        set index 0
        foreach exp $expected val $value {
            if { $exp != $val } {
                puts "Error: $msg (Values differ at index $index. Expected: $exp Have: $val)"
                return false
            }
            incr index
        }
        if { $printCheck } {
            puts "Check: $msg (List length. Expected: [llength $expected] Have: [llength $value])"
        }
        return true
    }

    proc CheckMatrix { expected value msg { printCheck true } } {
        # Check, if 2 matrices are identical.
        #
        # expected   - Expected matrix.
        # value      - Test matrix.
        # msg        - Message for test case.
        # printCheck - Print message for successful test case.
        #
        # Return true, if both matrices are identical.
        # If "printCheck" is set to true, a line prepended with "Check:" and the
        # message supplied in "msg" is printed to standard output.
        # If the check fails, return false and print message prepended with "Error:". 

        if { [llength $expected] != [llength $value] } {
            puts "Error: $msg (Matrix rows differ. Expected: [llength $expected] Have: [llength $value])"
            return false
        }
        set row 0
        foreach expRow $expected valRow $value {
            set col 0
            foreach exp $expRow val $valRow {
                if { $exp != $val } {
                    puts "Error: $msg (Values differ at row/col $row/$col. Expected: $exp Have: $val)"
                    return false
                }
                incr col
            }
            incr row
        }
        if { $printCheck } {
            puts "Check: $msg (Matrix rows. Expected: [llength $expected] Have: [llength $value])"
        }
        return true
    }
}

::Cawt::_Init
