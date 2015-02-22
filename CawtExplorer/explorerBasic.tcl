# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval ::Explorer {

    variable explorerAppName "InternetExplorer.Application"
    variable _ruffdoc

    lappend _ruffdoc Introduction {
        The Explorer namespace provides commands to control the Internet Explorer browser.

        Note: If running on Windows Vista or 7, you have to lower the
        security settings like follows:

        Internet Options -> Security -> Trusted Sites    : Low

        Internet Options -> Security -> Internet         : Medium + unchecked Enable Protected Mode

        Internet Options -> Security -> Restricted Sites : unchecked Enable Protected Mode
    }

    proc OpenNew { { visible true } { width -1 } { height -1 } } {
        # Open a new Internet Explorer instance.
        #
        # visible - true: Show the application window.
        #           false: Hide the application window.
        # width   - Width of the application window. If negative, open with last used width.
        # height  - Height of the application window. If negative, open with last used height.
        #
        # Return the identifier of the new Internet Explorer application instance.
        #
        # See also: Open Quit Visible

        variable explorerAppName

        set appId [::Cawt::GetOrCreateApp $explorerAppName false]
        ::Explorer::Visible $appId $visible
        if { $width >= 0 } {
            $appId Width [expr $width]
        }
        if { $height >= 0 } {
            $appId Height [expr $height]
        }
        return $appId
    }

    proc Open { { visible true } { width -1 } { height -1 } } {
        # Open an Internet Explorer instance. Use an already running instance, if available.
        # Use an already running Internet Explorer, if available.
        #
        # visible - true: Show the application window.
        #           false: Hide the application window.
        # width   - Width of the application window. If negative, open with last used width.
        # height  - Height of the application window. If negative, open with last used height.
        #
        # Return the identifier of the Internet Explorer application instance.
        #
        # See also: OpenNew Quit Visible

        variable explorerAppName

        set appId [::Cawt::GetOrCreateApp $explorerAppName true]
        ::Explorer::Visible $appId $visible
        if { $width >= 0 } {
            $appId Width [expr $width]
        }
        if { $height >= 0 } {
            $appId Height [expr $height]
        }
        return $appId
    }

    proc Quit { appId } {
        # Quit an Internet Explorer instance.
        #
        # appId - Identifier of the Internet Explorer instance.
        #
        # No return value.
        #
        # See also: Open

        $appId Quit
    }

    proc Visible { appId visible } {
        # Toggle the visibility of an Internet Explorer application window.
        #
        # appId   - Identifier of the Internet Explorer instance.
        # visible - true: Show the application window.
        #           false: Hide the application window.
        #
        # No return value.
        #
        # See also: Open OpenNew

        $appId Visible [::Cawt::TclInt $visible]
    }

    proc FullScreen { appId onOff } {
        # Toggle the fullscreen mode of an Internet Explorer application window.
        #
        # appId - Identifier of the Internet Explorer instance.
        # onOff - true: Use fullscreen mode.
        #         false: Use windowed mode.
        #
        # No return value.
        #
        # See also: Open Visible

        $appId FullScreen [::Cawt::TclBool $onOff]
    }

    proc Navigate { appId urlOrFile { wait true } { targetFrame "_self" } } {
        # Navigate to a URL or local file.
        #
        # appId       - Identifier of the Internet Explorer instance.
        # urlOrFile   - URL or local file name (as an absolute pathname).
        # wait        - Wait until page has been loaded completely.
        # targetFrame - Name of the frame in which to display the resource.
        #
        # The following predefined names for targetFrame are possible:
        # "_blank":  Load the link into a new unnamed window.
        # "_parent": Load the link into the immediate parent of the document the link is in.
        # "_self":   Load the link into the same window the link was clicked in.
        # "_top":    Load the link into the full body of the current window.
        #
        # If given any other string, it is interpreted as a named HTML frame.
        # If no frame or window exists that matches the specified target name,
        # a new window is opened for the specified link.
        #
        # No return value.
        #
        # See also: Open OpenNew

        $appId Navigate $urlOrFile 0 $targetFrame
        if { $wait } {
            while {[[$appId Document] readyState] != "complete"} {
                after 100
            }
        }
    }

    proc Go { appId target } {
        # Go to a specific page.
        #
        # appId  - Identifier of the Internet Explorer instance.
        # target - String identifying the target page.
        #
        # Possible values for target are: "Back", "Forward", "Home", "Search"
        #
        # No return value.

        set cmd "Go$target"
        eval $appId $cmd
    }

    proc IsBusy { appId } {
        # Check, if an Internet Explorer instance is busy.
        #
        # appId - Identifier of the Internet Explorer instance.
        #
        # Return true or false dependent on the busy status.

        return [expr {[$appId Busy]? true: false}]
    }
}
