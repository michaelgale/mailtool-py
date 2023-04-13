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
# MailAttachment class

from toolbox import safe_filename, split_filename, strip_rich_str, eng_units
from mailtool import *


class MailAttachment:
    def __init__(self, **kwargs):
        self.filename = ""
        self.payload = None
        self.size = 0
        for k, v in kwargs.items():
            if k in self.__dict__:
                self.__dict__[k] = v

    def __str__(self):
        s = self.colour_str()
        s = strip_rich_str(s)
        s = s.replace(":newspaper:", "[]")
        s = s.replace(":framed_picture: ", "[]")
        s = s.replace(":compression: ", "[]")
        return s

    def colour_str(self):
        s = []
        if self.is_document:
            s.append(":newspaper: ")
        if self.is_image:
            s.append(":framed_picture:  ")
        if self.is_archive:
            s.append(":compression:  ")
        s.append(colour_file_str(self.filename, " "))
        s.append(colour_size_str(self.size))
        return "".join(s)

    @property
    def ext(self):
        if self.filename is not None:
            _, e = split_filename(self.filename)
            return e.replace(".", "").lower()
        return None

    @property
    def is_archive(self):
        return any([self.ext == x for x in "zip rar 7z tar xz".split()])

    @property
    def is_image(self):
        return any([self.ext == x for x in "png jpg jpeg gif bmp".split()])

    @property
    def is_document(self):
        return any([self.ext == x for x in "docx doc pdf xls xlsx".split()])

    def write(self, fn=None):
        fn = fn if fn is not None else self.filename
        if self.payload is not None:
            with open(fn, "wb") as fp:
                fp.write(self.payload)

    @staticmethod
    def from_structure(s):
        a = MailAttachment()
        a.filename = MailAttachment.attach_filename(s)
        a.size = int(s[6])
        return a

    @staticmethod
    def attach_filename(s):
        def find_value(items, with_key):
            if items is None:
                return None
            item_count = len(items)
            for i, item in enumerate(items):
                if item.lower() == with_key and (i + 1) < item_count:
                    return str(items[i + 1], encoding="ascii")
            return None

        if len(s) < 10:
            return None
        fn = None
        for e in s:
            if isinstance(e, tuple):
                if e[0].lower() == b"attachment":
                    if e[1] is None:
                        fn = ""
                    else:
                        fn = find_value(e[1], b"filename")
                        fn = safe_filename(fn)
        if fn is not None and len(fn) == 0:
            for e in s:
                if isinstance(e, tuple):
                    if e[0].lower() == b"name":
                        fn = str(e[1], encoding="ascii")
                        fn = safe_filename(fn)
        if fn is not None and fn == "":
            return None
        return fn