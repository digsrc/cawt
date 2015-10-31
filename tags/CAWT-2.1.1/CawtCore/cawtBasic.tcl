# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval Cawt {

    namespace ensemble create

    namespace export CentiMetersToPoints
    namespace export Destroy
    namespace export GetComObjects
    namespace export GetDotsPerInch
    namespace export GetNumComObjects
    namespace export GetOrCreateApp
    namespace export GetPkgVersion
    namespace export HavePkg
    namespace export InchesToPoints
    namespace export IsComObject
    namespace export IsValidId
    namespace export KillApp
    namespace export PopComObjects
    namespace export PrintNumComObjects
    namespace export PushComObjects
    namespace export SetDotsPerInch
    namespace export TclBool
    namespace export TclInt
    namespace export TclString

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

        set retVal [catch {package require cawtoutlook} version]
        set pkgInfo(cawtoutlook,avail)   [expr !$retVal]
        set pkgInfo(cawtoutlook,version) $version

        set retVal [catch {package require cawtppt} version]
        set pkgInfo(cawtppt,avail)   [expr !$retVal]
        set pkgInfo(cawtppt,version) $version

        set retVal [catch {package require cawtword} version]
        set pkgInfo(cawtword,avail)   [expr !$retVal]
        set pkgInfo(cawtword,version) $version

        set retVal [catch {twapi::tclcast bstr "0.0"} val]
        set pkgInfo(haveStringCast) [expr !$retVal]
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
        # The version of the sub-package is returned as a string.
        # If the package is not available (HavePkg == false), an
        # empty string is returned.
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

    proc TclInt { val } {
        # Cast a value to an integer with boolean range.
        #
        # val - The value to be casted.
        #
        # Return 1, if val is not equal to zero or true.
        # Return 0, if val is equal to zero or false.
        #
        # See also: TclBool TclString

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
        # See also: TclInt TclString

        return [twapi::tclcast boolean $val]
    }

    proc TclString { val } {
        # Cast a value to a string.
        #
        # val - The value to be casted.
        #
        # Return casted string in a format usable for the COM interface.
        #
        # See also: TclInt TclBool

        variable pkgInfo

        if { $pkgInfo(haveStringCast) } {
            return [twapi::tclcast bstr $val]
        } else {
            return [twapi::tclcast string $val]
        }
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

    proc IsValidId { comObj } {
        # Obsolete: Replaced with IsComObject in version 2.0.0

        return [IsComObject $comObj]
    }

    proc IsComObject { comObj } {
        # Check, if a COM object is valid.
        #
        # comObj - The COM object.
        #
        # Return true, if comObj is a valid object.
        # Otherwise return false.
        #
        # See also: GetComObjects GetNumComObjects PrintNumComObjects

        return [expr { [twapi::comobj? $comObj] && ! [$comObj -isnull] } ]
    }

    proc GetComObjects {} {
        # Return the COM objects currently in use as a list.
        #
        # See also: IsComObject GetNumComObjects PrintNumComObjects Destroy

        return [twapi::comobj_instances]
    }

    proc GetNumComObjects {} {
        # Return the number of COM objects currently in use.
        #
        # See also: IsComObject GetComObjects PrintNumComObjects Destroy

        return [llength [GetComObjects]]
    }

    proc PrintNumComObjects {} {
        # Print the number of currently available COM objects to stdout.
        #
        # See also: IsComObject GetComObjects GetNumComObjects Destroy

        puts "Number of COM objects: [GetNumComObjects]"
    }

    proc _PrintComObjStack { msg } {
        variable comObjStack

        puts "$msg :"
        set num 1
        foreach entry $comObjStack {
            puts "$num: $entry"
            incr num
        }
    }

    proc PushComObjects { { printStack false } } {
        # Push current list of COM objects onto a stack.
        #
        # printStack - Print stack content after pushing onto stdout.
        #
        # See also: PopComObjects

        variable comObjStack

        lappend comObjStack [lsort -dictionary [GetComObjects]]

        if { $printStack } {
            _PrintComObjStack "PushComObjects"
        }
    }

    proc PopComObjects { { printStack false } } {
        # Pop last entry from COM objects stack.
        #
        # printStack - Print stack content after popping onto stdout.
        #
        # Pop last entry from COM objects stack and
        # remove all COM objects currently in use which
        # are not contained in the popped entry.
        #
        # See also: PushComObjects

        variable comObjStack

        set lastEntry [lindex $comObjStack end]
        set comObjStack [lrange $comObjStack 0 end-1]
        foreach comObj [lsort -dictionary [GetComObjects]] {
            if { [lsearch -exact $lastEntry $comObj] < 0 } {
                Cawt Destroy $comObj
            }
        }
        if { $printStack } {
            _PrintComObjStack "PopComObjects"
        }
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
        #
        # See also: PushComObjects PopComObjects

        if { $comObj ne "" } {
            $comObj -destroy
        } else {
            foreach obj [GetComObjects] {
                $obj -destroy
            }
        }
    }
}

Cawt::_Init
