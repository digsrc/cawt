    proc GetEnumTypes { } {
        # Return a list of available enumeration types.
        #
        # See also: GetEnumNames GetEnumVal GetEnum

        variable enums

        return [lsort [array names enums]]
    }

    proc GetEnumNames { enum } {
        # Return a list of names of a given enumeration type.
        #
        # See also: GetEnumTypes GetEnumVal GetEnum

        variable enums

        if { [info exists enums($enum)] } {
            foreach { key val } $enums($enum) {
                lappend nameList $key
            }
            return $nameList
        } else {
            return [list]
        }
    }

    proc GetEnumVal { enum } {
        # Return the numeric value of an enumeration name.
        #
        # See also: GetEnumTypes GetEnumNames GetEnum

        variable enums

        foreach enumType [GetEnumTypes] {
            set ind [lsearch -exact $enums($enumType) $enum]
            if { $ind >= 0 } {
                return [lindex $enums($enumType) [expr { $ind + 1 }]]
            }
        }
        return ""
    }

    proc GetEnum { enumOrString } {
        # Return the numeric value of an enumeration.
        #
        # See also: GetEnumTypes GetEnumVal GetEnumNames

        set retVal [catch { expr int($enumOrString) } enumInt]
        if { $retVal == 0 } {
            return $enumInt
        } else {
            return [GetEnumVal $enumOrString]
        }
    }
