# Copyright: 2007-2015 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval Excel {

    namespace ensemble create

    namespace export CsvRowToList
    namespace export CsvStringToMatrix
    namespace export GetCsvSeparatorChar
    namespace export ListToCsvRow
    namespace export MatrixToCsvString
    namespace export ReadCsvFile
    namespace export SetCsvSeparatorChar
    namespace export WriteCsvFile

    variable sSepChar

    proc _InitCsv {} {
        Excel SetCsvSeparatorChar
    }

    proc GetCsvSeparatorChar {} {
        # Return the column separator character.

        variable sSepChar

        return $sSepChar
    }

    proc SetCsvSeparatorChar { { separatorChar ";" } } {
        # Set the column separator character.
        #
        # separatorChar - The character used as the column separator.

        variable sSepChar

        set sSepChar $separatorChar
    }

    proc CsvRowToList { rowStr } {
        # Return a CSV encoded row as a list of column values.
        #
        # rowStr - CSV encoded row as string.
        #
        # See also: ListToCsvRow

        variable sSepChar

        set tmpList {}
        set wordCount 1
        set combine 0

        set wordList [split $rowStr $sSepChar]
        set floatSep [Excel GetFloatSeparator]

        foreach word $wordList {
            # TODO: Check conversion between different floating-point separators.
            if { 0 && $floatSep ne "." && \
                ([string is double $word] || [string is integer $word]) } {
                    puts "string map"
                set word [string map [list $floatSep "."] $word]
            }
            set len [string length $word]
            if { [string index $word end] eq "\"" } {
                set endQuote 1
            } else {
                set endQuote 0
            }
            if { [string index $word 0] eq "\"" } {
                set begQuote 1
            } else {
                set begQuote 0
            }

            if { $begQuote && $endQuote && ($len % 2 == 1) } {
                set onlyQuotes [regexp {^[\"]+$} $word]
                if { $onlyQuotes } {
                    if { $combine } {
                        set begQuote 0
                    } else {
                        set endQuote 0
                    }
                }
            }
            if { $begQuote && $endQuote && ($len == 2) } {
                set begQuote 0
                set endQuote 0
            }

            if { $begQuote && $endQuote } {
                lappend tmpList [string map {\"\" \"} [string range $word 1 end-1]]
                set combine 0
                incr wordCount
            } elseif { !$begQuote && $endQuote } {
                append tmpWord [string range $word 0 end-1]
                lappend tmpList [string map {\"\" \"} $tmpWord]
                set combine 0
                incr wordCount
            } elseif { $begQuote && !$endQuote } {
                set tmpWord [string range $word 1 end]
                append tmpWord $sSepChar
                set combine 1
            } else {
                if { $combine } {
                    append tmpWord  [string map {\"\" \"} $word]
                    append tmpWord $sSepChar
                } else {
                   lappend tmpList [string map {\"\" \"} $word]
                   set combine 0
                   incr wordCount
                }
            }
        }
        return $tmpList
    }

    proc ListToCsvRow { rowList } {
        # Return a list of column values as a CSV encoded row string.
        #
        # rowList - List of column values.
        #
        # See also: CsvRowToList

        variable sSepChar

        set rowStr ""
        set len1 [expr [llength $rowList] -1]
        set curVal 0
        set floatSep [Excel GetFloatSeparator]
        foreach val $rowList {
            set tmp [string map {\n\r \ } $val]
            if { [string first $sSepChar $tmp] >= 0 || \
                 [string first "\"" $tmp] >= 0 } {
                regsub -all {"} $tmp {""} tmp
                set tmp [format "\"%s\"" $tmp]
            }
            if { 0 && $floatSep ne "." && \
                ([string is double $tmp] || [string is integer $tmp]) } {
                    puts "string map"
                set tmp [string map [list "." $floatSep] $tmp]
            }
            if { $curVal < $len1 } {
                append rowStr $tmp $sSepChar
            } else {
                append rowStr $tmp
            }
            incr curVal
        }
        return $rowStr
    }

    proc MatrixToCsvString { matrixList } {
        # Return a CSV encoded table string from a matrix list.
        #
        # matrixList - Matrix with table data.
        #
        # See also: CsvStringToMatrix ListToCsvRow

        foreach rowList $matrixList {
            append str [Excel ListToCsvRow $rowList]
            append str "\n"
        }
        return [string range $str 0 end-1]
    }

    proc CsvStringToMatrix { csvString } {
        # Return a matrix from a CSV encoded table string.
        #
        # csvString - CSV encoded table as string.
        #
        # See also: MatrixToCsvString CsvRowToList

        set trimString [string trim $csvString '\0''\n']
        foreach row [lrange [split $trimString "\n"] 0 end] {
            set row [string trim $row "\r"]
            lappend matrixList [Excel CsvRowToList $row]
        }
        return $matrixList
    }

    proc ReadCsvFile { csvFileName { useHeader true } { numHeaderRows 0 } } {
        # Read a CSV table file into a matrix.
        #
        # csvFileName   - Name of the MediaWiki file.
        # useHeader     - true: Insert the header rows of the CSV file into the matrix.
        #                 false: Only transfer the table data.
        # numHeaderRows - Number of rows interpreted as header rows.
        #
        # Return the CSV table data as a matrix.
        # See SetMatrixValues for the description of a matrix representation.
        #
        # See also: WriteCsvFile

        variable sSepChar

        set matrixList {}
        set rowCount 1

        set catchVal [catch {open $csvFileName r} fp]
        if { $catchVal != 0 } {
            error "Could not open file \"$csvFileName\" for reading."
        }
        fconfigure $fp -translation crlf

        while { [gets $fp row] >= 0 } {
            if { $rowCount <= $numHeaderRows && ! $useHeader } {
                incr rowCount
                continue
            }

            set tmpList [Excel CsvRowToList $row]
            lappend matrixList $tmpList
            incr rowCount
        }

        close $fp
        return $matrixList
    }

    proc WriteCsvFile { matrixList csvFileName } {
        # Write the values of a matrix into a CSV file.
        #
        # matrixList  - Matrix with table data.
        # csvFileName - Name of the CSV file.
        #
        # See SetMatrixValues for the description of a matrix representation.
        #
        # No return value.
        #
        # See also: ReadCsvFile

        set catchVal [catch {open $csvFileName w} fp]
        if { $catchVal != 0 } {
            error "Could not open file \"$csvFileName\" for writing."
        }
        fconfigure $fp -translation binary

        foreach row $matrixList {
            puts -nonewline $fp [Excel ListToCsvRow $row]
            puts -nonewline $fp "\r\n"
        }
        close $fp
    }
}

Excel::_InitCsv
