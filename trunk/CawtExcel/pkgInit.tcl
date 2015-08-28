# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

proc _InitCawtExcel { dir version } {
    package provide cawtexcel $version

    source [file join $dir excelConst.tcl]
    source [file join $dir excelBasic.tcl]
    source [file join $dir excelUtil.tcl]
    source [file join $dir excelTablelist.tcl]
    source [file join $dir excelWord.tcl]
    source [file join $dir excelImgRaw.tcl]
    source [file join $dir excelMatlabFile.tcl]
    source [file join $dir excelMediaWiki.tcl]
    source [file join $dir excelWikit.tcl]
    source [file join $dir excelHtml.tcl]
    source [file join $dir excelChart.tcl]
    source [file join $dir excelCsv.tcl]
}
