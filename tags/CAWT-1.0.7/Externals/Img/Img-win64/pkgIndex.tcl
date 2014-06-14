package ifneeded zlibtcl 1.2.8 \
    [list load [file join $dir zlibtcl128.dll]]
package ifneeded pngtcl 1.4.12 \
    [list load [file join $dir pngtcl1412.dll]]
package ifneeded tifftcl 3.9.4 \
    [list load [file join $dir tifftcl394.dll]]
package ifneeded jpegtcl 8.4 \
    [list load [file join $dir jpegtcl84.dll]]
# -*- tcl -*- Tcl package index file
# --- --- --- Handcrafted, final generation by configure.
#
# $Id: pkgIndex.tcl.in 209 2009-03-03 17:19:05Z nijtmans $

package ifneeded img::base 1.4.2 [list load [file join $dir tkimg142.dll]]

# Compatibility hack. When asking for the old name of the package
# then load all format handlers and base libraries provided by tkImg.
# Actually we ask only for the format handlers, the required base
# packages will be loaded automatically through the usual package
# mechanism.

# When reading images without specifying it's format (option -format),
# the available formats are tried in reversed order as listed here.
# Therefore file formats with some "magic" identifier, which can be
# recognized safely, should be added at the end of this list.

package ifneeded Img 1.4.2 {
    package require img::window
    package require img::tga
    package require img::ico
    package require img::pcx
    package require img::sgi
    package require img::sun
    package require img::xbm
    package require img::xpm
    package require img::ps
    package require img::jpeg
    package require img::png
    package require img::tiff
    package require img::bmp
    package require img::ppm
    package require img::gif
    package require img::pixmap
    package provide Img 1.4.2
}

package ifneeded img::bmp 1.4.2 \
    [list load [file join $dir tkimgbmp142.dll]]
package ifneeded img::gif 1.4.2 \
    [list load [file join $dir tkimggif142.dll]]
package ifneeded img::ico 1.4.2 \
    [list load [file join $dir tkimgico142.dll]]
package ifneeded img::jpeg 1.4.2 \
    [list load [file join $dir tkimgjpeg142.dll]]
package ifneeded img::pcx 1.4.2 \
    [list load [file join $dir tkimgpcx142.dll]]
package ifneeded img::pixmap 1.4.2 \
    [list load [file join $dir tkimgpixmap142.dll]]
package ifneeded img::png 1.4.2 \
    [list load [file join $dir tkimgpng142.dll]]
package ifneeded img::ppm 1.4.2 \
    [list load [file join $dir tkimgppm142.dll]]
package ifneeded img::ps 1.4.2 \
    [list load [file join $dir tkimgps142.dll]]
package ifneeded img::sgi 1.4.2 \
    [list load [file join $dir tkimgsgi142.dll]]
package ifneeded img::sun 1.4.2 \
    [list load [file join $dir tkimgsun142.dll]]
package ifneeded img::tga 1.4.2 \
    [list load [file join $dir tkimgtga142.dll]]
package ifneeded img::tiff 1.4.2 \
    [list load [file join $dir tkimgtiff142.dll]]
package ifneeded img::window 1.4.2 \
    [list load [file join $dir tkimgwindow142.dll]]
package ifneeded img::xbm 1.4.2 \
    [list load [file join $dir tkimgxbm142.dll]]
package ifneeded img::xpm 1.4.2 \
    [list load [file join $dir tkimgxpm142.dll]]
package ifneeded img::dted 1.4.2 \
    [list load [file join $dir tkimgdted142.dll]]
package ifneeded img::raw 1.4.2 \
    [list load [file join $dir tkimgraw142.dll]]
