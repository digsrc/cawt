if {[catch {package present Tcl 8.6.3}]} return
package ifneeded Tk 8.6.3 [list load [file join $dir .. Tk tk86.dll] Tk]

