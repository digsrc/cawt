# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval Outlook {

    namespace ensemble create

    namespace export CreateMail
    namespace export GetVersion
    namespace export Open
    namespace export OpenNew
    namespace export Quit
    namespace export SendMail

    variable outlookVersion "0.0"
    variable outlookAppName "Outlook.Application"
    variable _ruffdoc

    lappend _ruffdoc Introduction {
        The Outlook namespace provides commands to control Microsoft Outlook.
    }

    proc GetVersion { appId { useString false } } {
        # Return the version of an Outlook application.
        #
        # appId     - Identifier of the Outlook instance.
        # useString - true: Return the version name (ex. "Outlook 2000").
        #             false: Return the version number (ex. "9.0").
        #
        # Both version name and version number are returned as strings.
        # Version number is in a format, so that it can be evaluated as a
        # floating point number.
        #
        # See also: GetCompatibilityMode GetExtString

        array set map {
            "7.0"  "Outlook 95"
            "8.0"  "Outlook 97"
            "9.0"  "Outlook 2000"
            "10.0" "Outlook 2002"
            "11.0" "Outlook 2003"
            "12.0" "Outlook 2007"
            "14.0" "Outlook 2010"
            "15.0" "Outlook 2013"
        }
        set versionString [$appId Version]
        set members [split $versionString "."]
        set version "[lindex $members 0].[lindex $members 1]"
        if { $useString } {
            if { [info exists map($version)] } {
                return $map($version)
            } else {
                return "Unknown Outlook version $version"
            }
        } else {
            return $version
        }
    }

    proc OpenNew { { width -1 } { height -1 } } {
        # Open a new Outlook instance.
        #
        # width   - Width of the application window. If negative, open with last used width.
        # height  - Height of the application window. If negative, open with last used height.
        #
        # Return the identifier of the new Outlook application instance.
        #
        # See also: Open Quit

        variable outlookAppName
	variable outlookVersion

        set appId [Cawt GetOrCreateApp $outlookAppName false]
        set outlookVersion [Outlook GetVersion $appId]
        if { $width >= 0 } {
            $appId Width [expr $width]
        }
        if { $height >= 0 } {
            $appId Height [expr $height]
        }
        return $appId
    }

    proc Open { { width -1 } { height -1 } } {
        # Open an Outlook instance. Use an already running instance, if available.
        #
        # width   - Width of the application window. If negative, open with last used width.
        # height  - Height of the application window. If negative, open with last used height.
        #
        # Return the identifier of the Outlook application instance.
        #
        # See also: OpenNew Quit

        variable outlookAppName
	variable outlookVersion

        set appId [Cawt GetOrCreateApp $outlookAppName true]
        set outlookVersion [Outlook GetVersion $appId]
        if { $width >= 0 } {
            $appId Width [expr $width]
        }
        if { $height >= 0 } {
            $appId Height [expr $height]
        }
        return $appId
    }

    proc Quit { appId { showAlert true } } {
        # Quit an Outlook instance.
        #
        # appId     - Identifier of the Outlook instance.
        # showAlert - true: Show an alert window, if there are unsaved changes.
        #             false: Quit without saving any changes.
        #
        # No return value.
        #
        # See also: Open OpenNew

        if { ! $showAlert } {
            Cawt ShowAlerts $appId false
        }
        $appId Quit
    }

    proc CreateMail { appId recipientList { subject "" } { body "" } { attachmentList {} } } {
        # Create a new Outlook mail.
        #
        # appId          - Identifier of the Outlook instance.
        # recipientList  - List of mail addresses.
        # subject        - Subject text.
        # body           - Mail body text.
        # attachmentList - List of files used as attachment.
        #
        # Return the identifier of the new mail object.
        #
        # See also: SendMail
 
        set mailId [$appId CreateItem $Outlook::olMailItem]

        $mailId Display
        foreach recipient $recipientList {
            $mailId -with { Recipients } Add $recipient
        }
        $mailId Body $body
        $mailId Subject $subject
        foreach attachment $attachmentList {
            $mailId -with { Attachments } Add [file nativename $attachment]
        }
        return $mailId
    }

    proc SendMail { mailId } {
        # Send an Outlook mail.
        #
        # mailId - Identifier of the Outlook mail object.
        #
        # No return value.
        #
        # See also: CreateMail

        $mailId Send
    }
}
