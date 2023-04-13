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
# MailAccount class

from collections import defaultdict
from datetime import date, timedelta, datetime

from rich import print

from toolbox import ymd_from_date_spec
from mailtool import *


def listify(obj):
    if not isinstance(obj, (list)):
        return [(obj)]
    return obj


class MailAccount:
    def __init__(self, account_name):
        self.name = account_name
        self.capbilities = []
        with MailSession(account_name=self.name) as session:
            self.capabilities = session.capabilities()
        self.logging = False
        self._folders = None

    def log(self, msg):
        if self.logging:
            print(msg)

    def has_capability(self, capability):
        if str.encode(capability.upper()) in self.capabilities:
            return True
        return False

    @property
    def has_move_capability(self):
        return self.has_capability("MOVE")

    @property
    def folders(self):
        if self._folders is None:
            with MailSession(account_name=self.name) as session:
                folders = session.list_folders()
                self._folders = [f[2] for f in folders]
        return self._folders

    @property
    def folder_count(self):
        return len(self.folders)

    def _get_message_ids(self, folder_name, unread=False):
        uids = []
        with MailSession(account_name=self.name) as session:
            if folder_name == "*":
                folders = self.folders
            else:
                folders = listify(folder_name)
            for folder in folders:
                session.select_folder(folder, readonly=True)
                if unread:
                    uids.extend(session.search(["UNSEEN"]))
                else:
                    uids.extend(session.search())
        return uids

    def get_all_message_ids(self, folder_name):
        return self._get_message_ids(folder_name=folder_name)

    def get_unread_ids(self, folder_name):
        return self._get_message_ids(folder_name=folder_name, unread=True)

    def _get_messages_with_specs(
        self,
        folder_name,
        spec_keys,
        spec_vals,
        unread_only=False,
    ):
        def _invert_vals(k, v):
            vl = []
            vs = []
            vs.append(k)
            vs.append(v)
            vl.append(vs)
            return vl

        uids = []
        with MailSession(account_name=self.name) as session:
            session.select_folder(folder_name, readonly=True)
            criteria = [] if not unread_only else ["UNSEEN"]
            specs = listify(spec_keys)
            vals = listify(spec_vals)
            for spec, val in zip(specs, vals):
                if "NOT" in spec:
                    keys = spec.split()
                    criteria.append(str.encode("NOT"))
                    criteria.append(_invert_vals(str.encode(keys[1]), val))
                elif "UNANSWERED" in spec:
                    criteria.append(str.encode("UNANSWERED"))
                else:
                    criteria.append(str.encode(spec))
                    criteria.append(val)
            uids.extend(session.search(criteria=criteria))
        return uids

    def get_messages_since_date(self, folder_name, since_date, unread_only=False):
        return self._get_messages_with_specs(
            folder_name, "SINCE", since_date, unread_only=unread_only
        )

    def get_messages_since_days_ago(self, folder_name, days_ago=1, unread_only=False):
        date_ago = datetime.datetime.now() - timedelta(days=days_ago)
        since_date = date(date_ago.year, date_ago.month, date_ago.day)
        return self._get_messages_with_specs(
            folder_name, "SINCE", since_date, unread_only=unread_only
        )

    def get_messages_before_date(self, folder_name, before_date, unread_only=False):
        return self._get_messages_with_specs(
            folder_name, "BEFORE", before_date, unread_only=unread_only
        )

    def get_messages_between_dates(
        self, folder_name, since_date, before_date, unread_only=False
    ):
        return self._get_messages_with_specs(
            folder_name,
            ["SINCE", "BEFORE"],
            [since_date, before_date],
            unread_only=unread_only,
        )

    def _get_messages_with_addresses(
        self, folder_name, addresses, fromto="FROM", unread_only=False, invert=False
    ):
        uids = []
        addresses = listify(addresses)
        for address in addresses:
            address_uids = self._get_messages_with_specs(
                folder_name,
                [fromto],
                address,
                unread_only=unread_only,
            )
            uids.extend(address_uids)
            self.log(
                "Found %d messages %s %s in %s for a total of %d"
                % (len(address_uids), fromto, address, folder_name, len(uids))
            )
        return uids

    def get_messages_from(self, folder_name, senders, unread_only=False):
        return self._get_messages_with_addresses(
            folder_name, senders, "FROM", unread_only=unread_only
        )

    def get_messages_to(self, folder_name, recipients, unread_only=False):
        return self._get_messages_with_addresses(
            folder_name, recipients, "TO", unread_only=unread_only
        )

    def get_messages(self, folder_name, **kwargs):
        """Get messages from mail folder with criteria in a dictionary such as:
        { "senders": "info@mail.com", "since": date(2000, 6, 30), "unread": True }
        valid keys: senders, recipients, since, before, unread, subject
        """
        uids = set()
        unread_only = False
        if "unread" in kwargs:
            unread_only = kwargs["unread"]
            if len(kwargs) == 1:
                new_uids = self.get_unread_ids(folder_name)
                uids.update(new_uids)
        for k, v in kwargs.items():
            new_uids = []
            if k == "senders":
                new_uids = self.get_messages_from(
                    folder_name, v, unread_only=unread_only
                )
            elif "not_senders" in k:
                new_uids = self._get_messages_with_specs(
                    folder_name, "NOT FROM", v, unread_only=unread_only
                )
            elif "recipients" in k:
                new_uids = self.get_messages_to(folder_name, v, unread_only=unread_only)
            elif "since" in k:
                new_uids = self.get_messages_since_date(
                    folder_name, v, unread_only=unread_only
                )
            elif "before" in k:
                new_uids = self.get_messages_before_date(
                    folder_name, v, unread_only=unread_only
                )
            elif "date" in k:
                d1, d2 = ymd_from_date_spec(v)
                new_uids = self.get_messages_between_dates(
                    folder_name, d1, d2, unread_only=unread_only
                )
            elif k == "subject":
                new_uids = self._get_messages_with_specs(
                    folder_name, "SUBJECT", v, unread_only=unread_only
                )
            elif "not_subject" in k:
                new_uids = self._get_messages_with_specs(
                    folder_name, "NOT SUBJECT", v, unread_only=unread_only
                )
            elif "not_replied" in k:
                new_uids = self._get_messages_with_specs(
                    folder_name, "UNANSWERED", "", unread_only=unread_only
                )
            else:
                continue
            if len(uids) > 0:
                uids = uids.intersection(new_uids)
            else:
                uids.update(new_uids)
            # print(
            #     "search key %s found %d items interesecting to %d items"
            #     % (k, len(new_uids), len(uids))
            # )
        uids = list(uids)
        for k, v in kwargs.items():
            if k == "attachments":
                uids = self.messages_with_attachments(uids, folder_name)
            elif k == "no_attachments":
                uids = self.messages_with_attachments(uids, folder_name, invert=True)
        return uids

    def move_messages(self, uids, from_folder, to_folder):
        if len(uids) < 1:
            return
        with MailSession(account_name=self.name) as session:
            session.select_folder(from_folder, readonly=False)
            if self.has_move_capability:
                session.move(uids, to_folder)
            else:
                self.log(
                    "Attempting to move %d messages from %s to %s"
                    % (len(uids), from_folder, to_folder)
                )
                session.copy(uids, to_folder)
                session.delete_messages(uids)
                session.expunge(uids)
            self.log(
                "%d messages moved from %s to %s" % (len(uids), from_folder, to_folder)
            )

    def delete_messages(self, uids, from_folder):
        if len(uids) < 1:
            return
        with MailSession(account_name=self.name) as session:
            self.log(
                "Attempting to delete %d messages from %s" % (len(uids), from_folder)
            )
            session.select_folder(from_folder, readonly=False)
            session.delete_messages(uids)
            session.expunge(uids)
            self.log("%d messages deleted from %s" % (len(uids), from_folder))

    def messages_with_attachments(self, uids, from_folder, invert=False):
        if uids is None or len(uids) < 1:
            return uids
        attach_uids = []
        with MailSession(account_name=self.name) as session:
            session.select_folder(from_folder, readonly=True)
            for uid, data in session.fetch(uids, FETCH_KEYS).items():
                msg = MailMessage.from_data(uid, data)
                if not msg.has_attachments and not invert:
                    continue
                if msg.has_attachments and invert:
                    continue
                attach_uids.append(uid)
        return attach_uids

    def download_attachments(self, uids, from_folder):
        if uids is None or len(uids) < 1:
            return
        with MailSession(account_name=self.name) as session:
            session.select_folder(from_folder, readonly=True)
            keys = FETCH_KEYS
            keys.append("BODY.PEEK[]")
            for msgid, data in session.fetch(uids, keys).items():
                msg = MailMessage.from_data(msgid, data)
                if not msg.has_attachments:
                    continue
                print(msg.colour_str(show_attachments=True))
                msg.download_attachment()

    def print_messages(
        self, uids, from_folder, show_id=False, show_attachments=False, show_name=False
    ):
        if uids is None or len(uids) < 1:
            return
        with MailSession(account_name=self.name) as session:
            session.select_folder(from_folder, readonly=True)
            keys = FETCH_KEYS
            if show_attachments:
                keys.append("BODY.PEEK[]")
            for msgid, data in session.fetch(uids, keys).items():
                msg = MailMessage.from_data(msgid, data)
                print(
                    msg.colour_str(
                        show_attachments=show_attachments, show_name=show_name
                    )
                )
            print("%d messages" % (len(uids)))

    def unique_senders(self, uids, from_folder, show_id=False):
        if uids is None or len(uids) < 1:
            return
        senders = defaultdict(int)
        with MailSession(account_name=self.name) as session:
            session.select_folder(from_folder, readonly=True)
            for msgid, data in session.fetch(uids, FETCH_KEYS).items():
                msg = MailMessage.from_data(msgid, data)
                senders[msg.address] += 1
        for k, v in sorted(senders.items(), key=lambda kv: kv[1], reverse=False):
            print("%3d messages from %s" % (v, k))
        print("%d messages" % (len(senders)))
