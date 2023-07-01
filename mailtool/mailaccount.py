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
import time

from rich import print

from toolbox import ymd_from_date_spec
from mailtool import *
from .helpers import ADDRESS_COLOUR


def listify(obj):
    if isinstance(obj, str):
        ls = obj.split()
        if len(ls) > 1:
            return ls
        return [obj]
    if not isinstance(obj, (list)):
        return [obj]
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

    def create_folder(self, folder_name):
        with MailSession(account_name=self.name) as session:
            session.create_folder(folder_name)
            self._folders = None

    def rename_folder(self, folder_name, new_name):
        with MailSession(account_name=self.name) as session:
            session.rename_folder(folder_name, new_name)
            self._folders = None

    def remove_folder(self, folder_name):
        with MailSession(account_name=self.name) as session:
            session.delete_folder(folder_name)
            self._folders = None

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
        uids = []
        with MailSession(account_name=self.name) as session:
            session.select_folder(folder_name, readonly=True)
            criteria = [] if not unread_only else ["UNSEEN"]
            specs = listify(spec_keys)
            vals = listify(spec_vals)
            for spec, val in zip(specs, vals):
                if "NOT" in spec:
                    criteria.append(str.encode("NOT"))
                    criteria.append([str.encode(specs[1]), val])
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

    def get_message_uids(self, folder_name, searchspec, merge="or"):
        """Get messages from mail folder with criteria in a dictionary such as:
        { "senders": "info@mail.com", "since": date(2000, 6, 30), "unread": True }
        valid keys: senders, recipients, since, before, unread, subject
        """
        uids = set()
        unread_only = False
        if "unread" in searchspec:
            unread_only = searchspec["unread"]
            if len(searchspec) == 1:
                new_uids = self.get_unread_ids(folder_name)
                uids.update(new_uids)
        elif len(searchspec) == 0:
            new_uids = self.get_all_message_ids(folder_name)
            uids.update(new_uids)
        first = True
        for k, v in searchspec.items():
            new_uids = []
            if k == "senders":
                if ">>" in v:
                    all_uids = self.get_all_message_ids(folder_name)
                    objs = self.get_message_objs(all_uids, folder_name)
                    for msg in objs:
                        ms = msg.address.split("@")
                        if len(ms) == 2:
                            if len(ms[1]) > 32:
                                new_uids.append(msg.uid)
                else:
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
                if merge == "and":
                    uids = uids.intersection(new_uids)
                else:
                    uids.update(new_uids)
            elif merge == "or" or len(searchspec) == 1:
                uids.update(new_uids)
            elif first and len(new_uids) > 0:
                uids.update(new_uids)
            # print(
            #     "search key %s(%s) found %d items interesecting (%s) to %d items"
            #     % (k, v, len(new_uids), merge, len(uids))
            # )
            first = False
        uids = list(uids)
        for k, v in searchspec.items():
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

    def get_message_objs(self, uids, from_folder, include_attach=False):
        if uids is None or len(uids) < 1:
            return
        objs = []
        with MailSession(account_name=self.name) as session:
            session.select_folder(from_folder, readonly=True)
            keys = FETCH_KEYS
            if include_attach:
                keys.append("BODY.PEEK[]")
            for msgid, data in session.fetch(uids, keys).items():
                msg = MailMessage.from_data(msgid, data)
                objs.append(msg)
        return objs

    def print_messages(self, msgs, show_attachments=False, show_name=False):
        if msgs is None:
            return
        for msg in msgs:
            print(
                msg.colour_str(show_attachments=show_attachments, show_name=show_name)
            )
        print("%d messages" % (len(msgs)))

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
            print("%3d messages from [%s]%s[/]" % (v, ADDRESS_COLOUR, k))
        print("%d messages" % (len(senders)))

    def show_rules(self):
        rulefile = MailSession.get_account_rules(self.name)
        if rulefile is None:
            return
        rules = metayaml.read(rulefile, disable_order_dict=True)
        print(rules)

    def process_rules(self, dryrun=False):
        rulefiles = MailSession.get_account_rules(self.name)
        if rulefiles is None:
            return
        for rulefile in rulefiles:
            print("  Processing rule file: %s" % (rulefile))
            rules = metayaml.read(rulefile, disable_order_dict=True)
            if not "rules" in rules:
                continue
            rules = rules["rules"]
            num_rules = len(rules)
            for i, rule in enumerate(rules):
                if "move" in rule:
                    print(
                        "  Rule:[white]%3d / %-3d : Move to [%s]%s"
                        % (i + 1, num_rules, FOLDER_COLOUR, rule["move"])
                    )
                elif "delete" in rule:
                    print("  Rule:[white]%3d / %-3d : Delete" % (i + 1, num_rules))
                search = {}
                for k, v in rule.items():
                    if any([k in SEARCH_KEYS]):
                        search[k] = listify(v)
                merge = "and"
                if "merge" in rule:
                    merge = rule["merge"]
                if len(search) > 0:
                    uids = self.get_message_uids("INBOX", search, merge=merge)
                    if len(uids) > 0:
                        msgs = self.get_message_objs(uids, "INBOX")
                        for msg in msgs:
                            print(msg.colour_str())
                        if "move" in rule:
                            if rule["move"] not in self.folders:
                                if not dryrun:
                                    print(
                                        "  [yellow bold]:warning:[/] Creating new folder [%s]%s[/] since it is not on the mail server"
                                        % (FOLDER_COLOUR, rule["move"])
                                    )
                                    self.create_folder(rule["move"])
                                else:
                                    print(
                                        "  [yellow bold]:warning:[/] Would create new folder [%s]%s[/] since it is not on the mail server"
                                        % (FOLDER_COLOUR, rule["move"])
                                    )
                            if not dryrun:
                                self.move_messages(uids, "INBOX", rule["move"])
                                # pace the IMAP actions with a sleep
                                time.sleep(1.5)
                                print("  Moved %d messages" % (len(uids)))
                            else:
                                print("  Would move %d messages" % (len(uids)))
                        if "delete" in rule:
                            if not dryrun:
                                self.delete_messages(uids, "INBOX")
                                # pace the IMAP actions with a sleep
                                time.sleep(1.5)
                                print("  Deleted %d messages" % (len(uids)))
                            else:
                                print("  Would delete %d messages" % (len(uids)))
