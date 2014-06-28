# Copyright: 2007-2014 Paul Obermeier (obermeier@poSoft.de)
# Distributed under BSD license.

namespace eval ::Cawt {

    proc CheckString { expected value msg { printCheck true } } {
        # Check, if 2 string values are identical.
        #
        # expected   - Expected string value.
        # value      - Test string value.
        # msg        - Message for test case.
        # printCheck - Print message for successful test case.
        #
        # Return true, if both string values are identical.
        # If "printCheck" is set to true, a line prepended with "Check:" and the
        # message supplied in "msg" is printed to standard output.
        # If the check fails, return false and print message prepended with "Error:".

        if { $expected ne $value } {
            puts "Error: $msg (Expected: \"$expected\" Have: \"$value\")"
            return false
        }
        if { $printCheck } {
            puts "Check: $msg (Expected: \"$expected\" Have: \"$value\")"
        }
        return true
    }

    proc CheckNumber { expected value msg { printCheck true } } {
        # Check, if 2 numerical values are identical.
        #
        # expected   - Expected numeric value.
        # value      - Test numeric value.
        # msg        - Message for test case.
        # printCheck - Print message for successful test case.
        #
        # Return true, if both numeric values are identical.
        # If "printCheck" is set to true, a line prepended with "Check:" and the
        # message supplied in "msg" is printed to standard output.
        # If the check fails, return false and print message prepended with "Error:".

        if { $expected != $value } {
            puts "Error: $msg (Expected: $expected Have: $value)"
            return false
        }
        if { $printCheck } {
            puts "Check: $msg (Expected: $expected Have: $value)"
        }
        return true
    }

    proc CheckList { expected value msg { printCheck true } } {
        # Check, if 2 lists are identical.
        #
        # expected   - Expected list.
        # value      - Test list.
        # msg        - Message for test case.
        # printCheck - Print message for successful test case.
        #
        # Return true, if both lists are identical.
        # If "printCheck" is set to true, a line prepended with "Check:" and the
        # message supplied in "msg" is printed to standard output.
        # If the check fails, return false and print message prepended with "Error:".

        if { [llength $expected] != [llength $value] } {
            puts "Error: $msg (List length differ. Expected: [llength $expected] Have: [llength $value])"
            return false
        }
        set index 0
        foreach exp $expected val $value {
            if { $exp != $val } {
                puts "Error: $msg (Values differ at index $index. Expected: $exp Have: $val)"
                return false
            }
            incr index
        }
        if { $printCheck } {
            puts "Check: $msg (List length. Expected: [llength $expected] Have: [llength $value])"
        }
        return true
    }

    proc CheckMatrix { expected value msg { printCheck true } } {
        # Check, if 2 matrices are identical.
        #
        # expected   - Expected matrix.
        # value      - Test matrix.
        # msg        - Message for test case.
        # printCheck - Print message for successful test case.
        #
        # Return true, if both matrices are identical.
        # If "printCheck" is set to true, a line prepended with "Check:" and the
        # message supplied in "msg" is printed to standard output.
        # If the check fails, return false and print message prepended with "Error:". 

        if { [llength $expected] != [llength $value] } {
            puts "Error: $msg (Matrix rows differ. Expected: [llength $expected] Have: [llength $value])"
            return false
        }
        set row 0
        foreach expRow $expected valRow $value {
            set col 0
            foreach exp $expRow val $valRow {
                if { $exp != $val } {
                    puts "Error: $msg (Values differ at row/col $row/$col. Expected: $exp Have: $val)"
                    return false
                }
                incr col
            }
            incr row
        }
        if { $printCheck } {
            puts "Check: $msg (Matrix rows. Expected: [llength $expected] Have: [llength $value])"
        }
        return true
    }
}
