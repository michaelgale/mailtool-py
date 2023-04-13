#! /usr/bin/env python3
#
# Copyright (C) 2023  Michael Gale

# Permission is hereby granted, free of charge, to any person
# obtaining a copy of this software and associated documentation
# files (the "Software"), to deal in the Software without restriction,
# including without limitation the rights to use, copy, modify, merge,
# publish, distribute, sublicense, and/or sell copies of the Software,
# and to permit persons to whom the Software is furnished to do so,
# subject to the following conditions:
#
# The above copyright notice and this permission notice shall be
# included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
# IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
# CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
# TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
# SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#
# Helper utility functions

from toolbox import rich_colour_str, eng_units

ACCOUNT_COLOUR = "#30C0A0"
FOLDER_COLOUR = "#F0D0A0"
UID_COLOUR = "#808080"
DATE_COLOUR = "#A090FF"
ADDRESS_COLOUR = "#6080FF"
FILE_COLOUR = "#F09070"
SUBJECT_COLOUR = "#A0A0A0"


def colour_date_str(date):
    return rich_colour_str(date, DATE_COLOUR)


def colour_address_str(s, suffix=""):
    return rich_colour_str(s, ADDRESS_COLOUR, bold=True, suffix=suffix)


def colour_file_str(s, suffix=""):
    return rich_colour_str(s, FILE_COLOUR, bold=True, suffix=suffix)


def colour_size_str(size, suffix=""):
    s = eng_units(size, units="B", sigfigs=4, unitary=True)
    return rich_colour_str(s, "white", bold=True, suffix=suffix)


def colour_subject_str(s, suffix=""):
    return rich_colour_str(s, SUBJECT_COLOUR, suffix=suffix)
