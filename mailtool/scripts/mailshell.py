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

import argparse
from datetime import date, timedelta, datetime

from rich import print

from toolbox import *
from mailtool import *


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-a", "--account", default="all", help="Mail account name")
    parser.add_argument(
        "-m", "--mailbox", default="INBOX", help="Mailbox folder name (default=INBOX)"
    )
    parser.add_argument(
        "-l",
        "--folders",
        action="store_true",
        default=False,
        help="List account folders",
    )
    parser.add_argument(
        "-us",
        "--unique",
        action="store_true",
        default=False,
        help="List unique senders",
    )
    parser.add_argument(
        "-t",
        "--tasks",
        action="store_true",
        default=False,
        help="Perform automated tasks",
    )
    parser.add_argument(
        "-o",
        "--output",
        action="store_true",
        default=False,
        help="Download message attachments",
    )
    parser.add_argument(
        "-u",
        "--unread",
        action="store_true",
        default=False,
        help="List unread messages",
    )
    parser.add_argument(
        "-n",
        "--name",
        action="store_true",
        default=False,
        help="Show sender name instead of address",
    )
    parser.add_argument(
        "-y", "--date", default=None, help="Date specifier as YYYY, YYYY-MM, YYYY-MM-DD"
    )
    parser.add_argument("-d", "--days", default=None, help="Days since today")
    parser.add_argument("-f", "--from", default=None, help="Messages sent from")
    parser.add_argument("-nf", "--notfrom", default=None, help="Messages not sent from")
    parser.add_argument("-s", "--subject", default=None, help="Subject")
    parser.add_argument("-ns", "--notsubject", default=None, help="Not in Subject")
    parser.add_argument(
        "-at",
        "--attachments",
        action="store_true",
        default=None,
        help="Has attachments",
    )
    parser.add_argument(
        "-nat",
        "--noattachments",
        action="store_true",
        default=None,
        help="No attachments",
    )
    parser.add_argument(
        "-nr",
        "--notreplied",
        action="store_true",
        default=False,
        help="List unreplied messages",
    )

    args = parser.parse_args()
    argsd = vars(args)

    search_dict = {}
    if argsd["days"]:
        date_ago = datetime.datetime.now() - timedelta(days=int(argsd["days"]))
        since_date = date(date_ago.year, date_ago.month, date_ago.day)
        search_dict["since"] = since_date
    if argsd["date"]:
        search_dict["date"] = argsd["date"]
    if argsd["from"]:
        search_dict["senders"] = argsd["from"]
    if argsd["notfrom"]:
        search_dict["not_senders"] = argsd["notfrom"]
    if argsd["subject"]:
        search_dict["subject"] = argsd["subject"]
    if argsd["notsubject"]:
        search_dict["not_subject"] = argsd["notsubject"]
    if argsd["unread"]:
        search_dict["unread"] = argsd["unread"]
    if argsd["notreplied"]:
        search_dict["not_replied"] = argsd["notreplied"]
    if argsd["attachments"] is not None:
        search_dict["attachments"] = argsd["attachments"]
    if argsd["noattachments"] is not None:
        search_dict["no_attachments"] = argsd["noattachments"]

    if argsd["account"] == "all":
        accounts = MailSession.get_accounts()
    else:
        accounts = [argsd["account"]]

    for account in accounts:
        if argsd["folders"]:
            m = MailAccount(account_name=account)
            for folder in m.folders:
                print(
                    "Mailbox : [#80FFC0 bold]%s[/] [white]/[/] [#20C040]%s[/]"
                    % (account, folder)
                )

        if len(search_dict) > 0:
            m = MailAccount(account_name=account)
            show_attach = any(
                [x in search_dict for x in ["attachments", "no_attachments"]]
            )
            if argsd["mailbox"] == "*":
                folders = m.folders
            else:
                folders = [argsd["mailbox"]]
            for folder in folders:
                print(
                    "Mailbox : [#80FFC0 bold]%s[/] [white]/[/] [#20C040]%s[/]"
                    % (account, folder)
                )
                uids = m.get_messages(folder, **search_dict)
                if argsd["unique"]:
                    m.unique_senders(uids, folder)
                elif argsd["output"]:
                    m.download_attachments(uids, folder)
                else:
                    m.print_messages(
                        uids,
                        folder,
                        show_attachments=show_attach,
                        show_name=argsd["name"],
                    )

if __name__ == "__main__":
    main()

