# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval Cawt {

    namespace ensemble create

    namespace export ColorToRgb
    namespace export GetActivePrinter
    namespace export GetApplicationId
    namespace export GetApplicationName
    namespace export GetApplicationVersion
    namespace export GetDocumentProperties
    namespace export GetDocumentProperty
    namespace export GetInstallationPath
    namespace export GetStartupPath
    namespace export GetTemplatesPath
    namespace export GetUserLibraryPath
    namespace export GetUserName
    namespace export GetUserPath
    namespace export IsApplicationId
    namespace export RgbToColor
    namespace export SetDocumentProperty
    namespace export SetPrinterCommunication
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

    proc SetPrinterCommunication { objId onOff } {
        # Enable or disable printer communication.
        #
        # objId - The identifier of an Office object.
        # onOff - true : Printer communication is enabled.
        #         false: Printer communication is disabled.
        #
        # Disable the printer communication to speed up the execution of code
        # that sets PageSetup properties, ex. SetWorksheetPrintOptions.
        # Enable the printer communication after setting properties to commit
        # all cached PageSetup commands.
        #
        # No return value.
        #
        # See also: GetActivePrinter

        if { ! [Cawt IsApplicationId $objId] } {
            set appId [Cawt GetApplicationId $objId]
            $appId PrintCommunication [Cawt TclBool $onOff]
            Cawt Destroy $appId
        } else {
            $objId PrintCommunication [Cawt TclBool $onOff]
        }
    }

    proc GetActivePrinter { appId } {
        # Get the name of the active printer.
        #
        # appId - The application identifier.
        #
        # Return the name of the active printer as a string.
        #
        # See also: SetPrinterCommunication

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

    proc GetDocumentProperties { objId { type "" } } {
        # Get document property names as a list.
        #
        # objId - The identifier of an Office object (Workbook, Document, Presentation).
        # type  - Type of document properties ("Builtin" or "Custom").
        #         If type is not specified or the empty string, both types
        #         of document properties are included in the list.
        #
        # Return a sorted Tcl list containing the names of all properties
        # of the specified type.
        #
        # See also: GetDocumentProperty SetDocumentProperty

        set propsBuiltin [$objId BuiltinDocumentProperties]
        set propsCustom  [$objId CustomDocumentProperties]

        set propList [list]
        if { $type eq "Builtin" || $type eq "" } {
            $propsBuiltin -iterate prop {
                lappend propList [$prop Name]
                Cawt Destroy $prop
            }
        }
        if { $type eq "Custom" || $type eq "" } {
            $propsCustom -iterate prop {
                lappend propList [$prop Name]
                Cawt Destroy $prop
            }
        }
        Cawt Destroy $propsBuiltin
        Cawt Destroy $propsCustom
        return [lsort -dictionary $propList]
    }

    proc _GetPropertyValue { propertyId } {
        set retVal [catch {$propertyId Value} propVal]
        if { $retVal == 0 } {
            return $propVal
        } else {
            return "N/A"
        }
    }

    proc GetDocumentProperty { objId propertyName } {
        # Get the value of a document property.
        #
        # objId        - The identifier of an Office object (Workbook, Document, Presentation).
        # propertyName - The name of the property.
        #
        # Return the value of specified property.
        # If the property value is not set or an invalid property name is given,
        # the string "N/A" is returned.
        #
        # See also: GetDocumentProperties SetDocumentProperty

        set properties [Cawt GetDocumentProperties $objId]
        if { [lsearch $properties $propertyName] >= 0 } {
            set propsBuiltin [$objId BuiltinDocumentProperties]
            set retVal [catch {$propsBuiltin -get Item $propertyName} property]
            Cawt Destroy $propsBuiltin
            if { $retVal != 0 } {
                set propsCustom  [$objId CustomDocumentProperties]
                set retVal [catch {$propsCustom -get Item $propertyName} property]
                Cawt Destroy $propsCustom
                if { $retVal != 0 } {
                    set propertyValue "N/A"
                } else {
                    set propertyValue [_GetPropertyValue $property]
                    Cawt Destroy $property
                }
            } else {
                set propertyValue [_GetPropertyValue $property]
                Cawt Destroy $property
            }
        } else {
            error "GetDocumentProperty: \"$propertyName\" is not a valid property name."
        }
        return $propertyValue
    }

    proc SetDocumentProperty { objId propertyName propertyValue } {
        # Set the value of a document property.
        #
        # objId         - The identifier of an Office object (Workbook, Document, Presentation).
        # propertyName  - The name of the property to set.
        # propertyValue - The value for the property.
        #
        # No return value.
        #
        # If the property name is a builtin property, it's value is set.
        # Ptherwise a new custom property is generated and it's value set.
        #
        # See also: GetDocumentProperties GetDocumentProperty

        set properties [Cawt GetDocumentProperties $objId "Builtin"]
        if { [lsearch $properties $propertyName] >= 0 } {
            set propsBuiltin [$objId BuiltinDocumentProperties]
            $propsBuiltin -set Item $propertyName $propertyValue
            Cawt Destroy $propsBuiltin
        } else {
            set propsCustom [$objId CustomDocumentProperties]
            $propsCustom Add $propertyName [Cawt TclBool false] 4 $propertyValue
            Cawt Destroy $propsCustom
        }
    }
}
