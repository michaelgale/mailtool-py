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
# MailMessage class

import email
import mimetypes

from toolbox import str_from_mime_words, safe_filename, strip_rich_str
from mailtool import *


class MailMessage:
    def __init__(self, **kwargs):
        self.address = ""
        self.name = ""
        self.uid = None
        self.seq = None
        self.date = None
        self.subject = ""
        self.attachments = []
        self.ignore_file_ext = ["txt", "html", "gif"]
        self.include_file_ext = None
        self.seen = False
        self.answered = False
        self.flagged = False
        self.draft = False
        self.deleted = False
        for k, v in kwargs.items():
            if k in self.__dict__:
                self.__dict__[k] = v

    @property
    def has_attachments(self):
        return self.attachment_count > 0

    @property
    def attachment_count(self):
        return len(self.attachments)

    def get_attachment(self, filename):
        if self.has_attachments:
            for f in self.attachments:
                if f.filename == filename:
                    return f
        return None

    def __str__(self):
        s = self.colour_str()
        s = strip_rich_str(s)
        s = s.replace(":blue_circle:", "*")
        s = s.replace(":paperclip:", " &")
        s = s.replace(":file_folder:", "[]")
        s = s.replace(":leftwards_arrow_with_hook:", "<")
        s = s.replace(":triangular_flag:", ">")
        return s

    def colour_str(self, show_attachments=False, show_name=False):
        s = []
        if self.seen:
            s.append("   [#808080]%5s[/] " % (str(self.uid)))
        else:
            s.append(":blue_circle: [#FFFFFF]%5s[/] " % (str(self.uid)))
        if not self.answered:
            s.append("  ")
        else:
            s.append(":leftwards_arrow_with_hook: ")
        if self.flagged:
            s.append(":triangular_flag: ")
        s.append(colour_date_str(self.date))
        s.append("%2s " % (":paperclip:" if self.has_attachments else ""))
        maxlen = 72
        if show_name and len(self.name) > 0:
            name = self.name[:40]
            s.append(colour_address_str(name, " : "))
            maxlen -= len(name)
        else:
            addr = self.address[:40]
            s.append(colour_address_str(addr, " : "))
            maxlen -= len(addr)
        s.append(colour_subject_str(self.subject[:maxlen], " "))
        if self.has_attachments:
            for _ in range(self.attachment_count):
                s.append(":file_folder: ")
            total_size = sum([a.size for a in self.attachments])
            s.append(colour_size_str(total_size))
        if show_attachments and self.attachment_count > 0:
            s.append("\n")
            for a in self.attachments:
                s.append("    %s\n" % (a.colour_str()))
        return "".join(s).rstrip()

    def _data_from_dict(self, data, key):
        if isinstance(data, (dict)):
            if key in data:
                return data[key]
        return None

    def process_flags(self, data):
        flags = self._data_from_dict(data, key=b"FLAGS")
        if flags is None:
            return
        for f in flags:
            fs = str(f, encoding="ascii")
            fs = fs.replace("\\", "").lower()
            if fs == "seen":
                self.seen = True
            elif fs == "answered":
                self.answered = True
            elif fs == "flagged":
                self.flagged = True
            elif fs == "draft":
                self.draft = True
            elif fs == "deleted":
                self.deleted = True

    def process_envelope(self, data):
        e = self._data_from_dict(data, key=b"ENVELOPE")
        if e is None:
            return
        self.address = "%s@%s" % (
            str_from_mime_words(e.from_[0].mailbox),
            str_from_mime_words(e.from_[0].host),
        )
        if e.from_[0].name is not None:
            self.name = str_from_mime_words(e.from_[0].name)

        self.subject = str_from_mime_words(e.subject)
        self.date = e.date

    def walk_structure(self, msg, items=None):
        if items is None:
            items = []
        if isinstance(msg, (list, tuple)):
            if msg[0] is not None and not isinstance(msg[0], (list, tuple)):
                if any([x in msg[0].lower() for x in [b"application", b"image"]]):
                    if MailAttachment.attach_filename(msg) is not None:
                        items.append(msg)
            for e in msg:
                items = self.walk_structure(e, items=items)
        return items

    def process_structure(self, data):
        s = self._data_from_dict(data, key=b"BODYSTRUCTURE")
        if s is None:
            return
        attachments = self.walk_structure(s)
        self.attachments = []
        if attachments is not None and len(attachments) > 0:
            for attachment in attachments:
                a = MailAttachment.from_structure(attachment)
                self.attachments.append(a)

    def process_body(self, data, get_payload=True):
        d = self._data_from_dict(data, key=b"BODY[]")
        if d is None:
            return
        msg = email.message_from_bytes(d)
        counter = 1
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            filename = safe_filename(part.get_filename())
            if not filename:
                ext = mimetypes.guess_extension(part.get_content_type())
                if not ext:
                    ext = ".bin"
                filename = f"part-{counter:03d}{ext}"
            counter += 1
            if not any([filename.endswith(x) for x in self.ignore_file_ext]):
                if self.include_file_ext is not None:
                    if not any([filename.endswith(x) for x in self.include_file_ext]):
                        continue
                a = self.get_attachment(filename)
                if a is not None and get_payload:
                    a.payload = part.get_payload(decode=True)

    def download_attachment(self, file=None, new_name=None, limit=0):
        if not self.has_attachments or self.attachments is None:
            return
        if file is None:
            for i, f in enumerate(self.attachments):
                if limit is not None and limit > 0:
                    if not i < limit:
                        break
                f.write()
        else:
            for f in self.attachments:
                if file == f.filename:
                    if new_name is not None:
                        f.write(new_name)
                    else:
                        f.write()

    @staticmethod
    def from_data(uid, data):
        msg = MailMessage(uid=uid)
        if b"ENVELOPE" in data:
            msg.process_envelope(data)
        if b"FLAGS" in data:
            msg.process_flags(data)
        if b"BODYSTRUCTURE" in data:
            msg.process_structure(data)
        if b"BODY[]" in data:
            msg.process_body(data)
        return msg
