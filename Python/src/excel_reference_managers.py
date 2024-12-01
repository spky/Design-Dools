import re
import math

def colon_reference_to_int(colon_reference):
    """Splits a string of format [A-Z][1-9]:[A-Z][1-9], converts them to 
    integer lists, and returns a 2 element list of those lists"""
    references = colon_reference.split(":")
    if len(references) == 2:
        left_top = reference_to_int(references[0])
        right_bottom = reference_to_int(references[1])
    else:
        raise ValueError("Reference not in a format of (ref):(ref)")
    return [left_top, right_bottom]

def reference_to_int(reference):
    """Returns a list of [column, row], indexed at 0"""
    parsed = parse_reference(reference)
    if parsed is not None:
        column = column_to_int(parsed[0])
        row = int(parsed[1]) - 1
        return [column, row]
    else:
        return None

def parse_reference(reference):
    """Checks if the string matches a regex for an excel reference. 
    Returns a list consisting of the column and row references if it does 
    match. If it doesn't match it just returns None"""
    regex = re.compile("^(?P<column>[a-zA-Z]+)(?P<row>[0-9]+)$")
    match = re.match(regex, reference)
    if match is not None:
        return [match["column"], match["row"]]
    else:
        return None

def int_to_reference(int_reference_list):
    """Returns the equivalent reference to integer defined, indexed-0 column and row
    return int_to_column(column) + str(row)"""
    if int_reference_list[1] < 0:
        raise ValueError("Row cannot be less than 1")
    return int_to_column(int_reference_list[0]) + str(int_reference_list[1]+1)

def column_to_int(column_reference):
    """Returns column number corresponding to letter reference, ex: a or A = 0, 
    AA = 26. Letter case type is ignored.
    """
    base = 0
    integer = -1
    for letter in reversed(list(column_reference)):
        lower_letter = letter.lower()
        integer += (ord(lower_letter) - 96)*(26**base)
        base += 1
    return integer

def int_to_column(integer):
    """Returns the column reference letter, ex: 0 = A, 26 == AA. Always returns 
    capital letters"""
    # The alphabetical column system is a base-26 number system
    BASE = 26
    ASCII_OFFSET = 96
    column = ""
    # Initialize int_quotient and modulus so the loop will run at least once
    int_quotient = -1
    modulus = 0
    while integer:
        int_quotient = math.floor(integer/BASE)
        modulus = integer % BASE
        if int_quotient:
            column += chr(int_quotient + ASCII_OFFSET).upper()
            integer = modulus
        else:
            integer = 0
    column += chr(modulus + ASCII_OFFSET + 1).upper()
    return column