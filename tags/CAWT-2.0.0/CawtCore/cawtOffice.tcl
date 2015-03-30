# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval Cawt {

    namespace ensemble create

    namespace export ColorToRgb
    namespace export GetActivePrinter
    namespace export GetApplicationId
    namespace export GetApplicationName
    namespace export GetApplicationVersion
    namespace export GetInstallationPath
    namespace export GetStartupPath
    namespace export GetTemplatesPath
    namespace export GetUserLibraryPath
    namespace export GetUserName
    namespace export GetUserPath
    namespace export IsApplicationId
    namespace export RgbToColor
    namespace export ShowAlerts

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
        #
        # See also: ColorToRgb

        return [expr {int ($b) << 16 | int ($g) << 8 | int($r)}]
    }

    proc ColorToRgb { color } {
        # Return an Office color number as a RGB color list.
        #
        # color - The Office color number
        #
        # The r, g and b values are returned as integers in the
        # range 0 .. 255.
        #
        # Return the color number as a list of r, b and b values.
        #
        # See also: RgbToColor

        set r [expr { (int ($color))       & 0xFF }]
        set g [expr { (int ($color) >>  8) & 0xFF }]
        set b [expr { (int ($color) >> 16) & 0xFF }]
        return [list $r $g $b]
    }

    proc ShowAlerts { appId onOff } {
        # Toggle the display of Office alerts.
        #
        # appId - The application identifier.
        # onOff - Switch the alerts on or off.
        #
        # No return value.

        if { $onOff } {
            if { [Cawt GetApplicationName $appId] eq "Microsoft Word" } {
                set alertLevel [expr $Word::wdAlertsAll]
            } else {
                set alertLevel [expr 1]
            }
        } else {
            set alertLevel [expr 0]
        }
        $appId DisplayAlerts $alertLevel
    }

    proc IsApplicationId { objId } {
        # Check, if Office object is an application identifier.
        #
        # objId - The identifier of an Office object.
        #
        # Return true
        # Return true, if objId is a valid Office application identifier.
        # Otherwise return false.
        #
        # See also: IsComObj GetApplicationId GetApplicationName

        set retVal [catch {$objId Version} errMsg]
        # Version is a property of all Office application classes.
        if { $retVal == 0 } {
            return true
        } else {
            return false
        }
    }

    proc GetApplicationId { objId } {
        # Get the application identifier of an Office object.
        #
        # objId - The identifier of an Office object.
        #
        # Office object are Workbooks, Worksheets, ...
        #
        # See also: GetApplicationName IsApplicationId

        return [$objId Application]
    }

    proc GetApplicationName { objId } {
        # Get the name of an Office application.
        #
        # objId - The identifier of an Office object.
        #
        # Return the name of the application as a string.
        #
        # See also: GetApplicationId IsApplicationId

        if { ! [Cawt IsApplicationId $objId] } {
            set appId [Cawt GetApplicationId $objId]
            set name [$appId Name]
            Cawt Destroy $appId
            return $name
        } else {
            return [$objId Name]
        }
    }

    proc GetApplicationVersion { objId } {
        # Get the version number of an Office application.
        #
        # objId - The identifier of an Office object.
        #
        # Return the version of the application as a floating point number.
        #
        # See also: GetApplicationId GetApplicationName

        if { ! [Cawt IsApplicationId $objId] } {
            set appId [Cawt GetApplicationId $objId]
            set version [$appId Version]
            Cawt Destroy $appId
        } else {
            set version [$objId Version]
        }
        return $version
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
}
