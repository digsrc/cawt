# Copyright: 2011-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval ::Matlab {

    variable matlabAppName "Matlab.Application"
    variable _ruffdoc

    lappend _ruffdoc Introduction {
        The Matlab namespace provides commands to control MathWorks Matlab.
    }

    proc OpenNew { { visible true } } {
        # Open a new Matlab instance.
        #
        # visible - true: Show the application window.
        #           false: Hide the application window.
        #
        # Return the identifier of the new Matlab application instance.
        #
        # See also: Open Visible Quit

        variable matlabAppName

        set appId [::Cawt::GetOrCreateApp $matlabAppName false]
        ::Matlab::Visible $appId $visible
        return $appId
    }

    proc Open { { visible true } } {
        # Open a Matlab instance. Use an already running Matlab, if available.
        #
        # visible - true: Show the application window.
        #           false: Hide the application window.
        #
        # Return the identifier of the Matlab application instance.
        #
        # See also: OpenNew Visible Quit

        variable matlabAppName

        set appId [::Cawt::GetOrCreateApp $matlabAppName true]
        ::Matlab::Visible $appId $visible
        return $appId
    }

    proc Visible { appId visible } {
        # Toggle the visibilty of a Matlab application window.
        #
        # appId   - Identifier of the Matlab instance.
        # visible - true: Show the application window.
        #           false: Hide the application window.
        #
        # No return value.
        #
        # See also: Open OpenNew

        $appId Visible [::Cawt::TclInt $visible]
    }

    proc Quit { appId } {
        # Quit a Matlab application instance.
        #
        # appId - Identifier of the Matlab instance.
        #
        # No return value.
        #
        # See also: Open

        $appId Quit
    }

    proc ExecCmd { appId cmd } {
        # Execute a Matlab command.
        #
        # appId - Identifier of the Matlab instance.
        # cmd   - String containg the Matlab command being executed.
        #
        # Return the Matlab answer as a string.

        set retVal [$appId Execute $cmd]
        return $retVal
    }
}
