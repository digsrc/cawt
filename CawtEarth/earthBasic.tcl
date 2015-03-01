# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.
# Idea taken from http://wiki.tcl.tk/24099

namespace eval Earth {

    namespace ensemble create

    namespace export IsInitialized
    namespace export Open
    namespace export OpenNew
    namespace export Quit
    namespace export SaveImage
    namespace export SetCamera

    variable earthAppName  "GoogleEarth.ApplicationGE"
    variable earthProgName "GoogleEarth.exe"
    variable _ruffdoc

    lappend _ruffdoc Introduction {
        The Earth namespace provides commands to control Google Earth.
    }

    proc OpenNew {} {
        # Open a new GoogleEarth instance.
        #
        # Return the identifier of the new GoogleEarth application instance.
        #
        # See also: Open Quit

        variable earthAppName

        set appId [Cawt GetOrCreateApp $earthAppName false]
        return $appId
    }

    proc Open {} {
        # Open a GoogleEarth instance. Use an already running instance, if available.
        #
        # Return the identifier of the GoogleEarth application instance.
        #
        # See also: OpenNew Quit

        variable earthAppName

        set appId [Cawt GetOrCreateApp $earthAppName true]
        return $appId
    }

    proc Quit { appId } {
        # Quit a GoogleEarth instance.
        #
        # appId - Identifier of the GoogleEarth instance.
        #
        # No return value.
        #
        # See also: Open

        variable earthProgName

        # Quit method not available, so we kill the application.
        # This may have some strnage effects in comination with MediaPlayer.
        # Cawt KillApp $earthProgName
    }

    proc IsInitialized { appId } {
        # Check, if a GoogleEarth instance is initialized.
        #
        # appId - Identifier of the GoogleEarth instance.
        #
        # Return true or false dependent on the initialization status.

        return [$appId IsInitialized]
    }

    proc SetCamera { appId latitude longitude altitude elevation azimuth } {
        # Set camera position and orientation.
        #
        # appId     - Identifier of the GoogleEarth instance.
        # latitude  - Latitude in degrees.  Range [-90.0, 90.0].
        # longitude - Longitude in degrees. Range [-180.0, 180.0].
        # altitude  - Altitude in meters.
        # elevation - Elevation angle in degrees. Range [0.0, 90.0].
        #              0 degrees corresponds looking to the center of the earth.
        #             90 degrees corresponds looking to the horizon.
        # azimuth   - Azimuth angle in degrees. Range [0.0, 360.0].
        #              0 degrees corresponds looking north.
        #             90 degrees corresponds loooking east.
        #
        # No return value.

        # SetCameraParams <latitude> <longitude> <altitude> <altMode> <range>
        #                 <tilt> <azimuth> <speed>
        #
        # <latitude>  Latitude in degrees.
        # <longitude> Longitude in degrees.
        # <altitude>  Altitude in meters.
        # <altMode>   Altitude mode that defines altitude reference origin.
        #             (1=above ground, 2=absolute)
        # <range>     Distance between focus point and camera in meters.
        #             If !=0 camera will move backward from range meters along the camera axis
        # <tilt>      Tilt angle in degrees.
        # <azimuth>   Azimuth angle in degrees.
        # <speed>     Speed factor. Must be >= 0, if >=5.0 it's the teleportation mode

        $appId SetCameraParams \
               $latitude $longitude 0.0 1 $altitude \
               $elevation $azimuth 6.0
    }

    proc SaveImage { appId fileName { quality 80 } } {
        # Save a grey-scale image of the current view.
        #
        # appId    - Identifier of the GoogleEarth instance.
        # fileName - Name of image file.
        # quality  - Quality of the JPEG compression in percent.
        #
        # No return value.

        $appId SaveScreenShot [file nativename $fileName] $quality
    }
}
