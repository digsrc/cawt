# Test mail functionality of the CawtOutlook package.
#
# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

source "SetTestPathes.tcl"
package require cawt

set appId [Outlook OpenNew]

set toList [list  \
    "info@poSoft.de" \
]

set attachmentList [list \
    [file nativename [file join [pwd] Outlook-02_Mail.tcl]] \
]

set mailId [Outlook CreateMail $appId $toList "Subject" "Body text line 1.\nBody text line 2." $attachmentList]
Outlook SendMail $mailId

if { [lindex $argv 0] eq "auto" } {
    Outlook Quit $appId
    ::Cawt::Destroy
    exit 0
}
::Cawt::Destroy
