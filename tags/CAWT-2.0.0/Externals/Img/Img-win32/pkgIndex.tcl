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

package ifneeded img::base 1.4.3 [list load [file join $dir tkimg143.dll]]

# Compatibility hack. When asking for the old name of the package
# then load all format handlers and base libraries provided by tkImg.
# Actually we ask only for the format handlers, the required base
# packages will be loaded automatically through the usual package
# mechanism.

# When reading images without specifying it's format (option -format),
# the available formats are tried in reversed order as listed here.
# Therefore file formats with some "magic" identifier, which can be
# recognized safely, should be added at the end of this list.

package ifneeded Img 1.4.3 {
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
    package provide Img 1.4.3
}

package ifneeded img::bmp 1.4.3 \
    [list load [file join $dir tkimgbmp143.dll]]
package ifneeded img::gif 1.4.3 \
    [list load [file join $dir tkimggif143.dll]]
package ifneeded img::ico 1.4.3 \
    [list load [file join $dir tkimgico143.dll]]
package ifneeded img::jpeg 1.4.3 \
    [list load [file join $dir tkimgjpeg143.dll]]
package ifneeded img::pcx 1.4.3 \
    [list load [file join $dir tkimgpcx143.dll]]
package ifneeded img::pixmap 1.4.3 \
    [list load [file join $dir tkimgpixmap143.dll]]
package ifneeded img::png 1.4.3 \
    [list load [file join $dir tkimgpng143.dll]]
package ifneeded img::ppm 1.4.3 \
    [list load [file join $dir tkimgppm143.dll]]
package ifneeded img::ps 1.4.3 \
    [list load [file join $dir tkimgps143.dll]]
package ifneeded img::sgi 1.4.3 \
    [list load [file join $dir tkimgsgi143.dll]]
package ifneeded img::sun 1.4.3 \
    [list load [file join $dir tkimgsun143.dll]]
package ifneeded img::tga 1.4.3 \
    [list load [file join $dir tkimgtga143.dll]]
package ifneeded img::tiff 1.4.3 \
    [list load [file join $dir tkimgtiff143.dll]]
package ifneeded img::window 1.4.3 \
    [list load [file join $dir tkimgwindow143.dll]]
package ifneeded img::xbm 1.4.3 \
    [list load [file join $dir tkimgxbm143.dll]]
package ifneeded img::xpm 1.4.3 \
    [list load [file join $dir tkimgxpm143.dll]]
package ifneeded img::dted 1.4.3 \
    [list load [file join $dir tkimgdted143.dll]]
package ifneeded img::raw 1.4.3 \
    [list load [file join $dir tkimgraw143.dll]]
