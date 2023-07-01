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
from argparse import RawDescriptionHelpFormatter
from datetime import date, timedelta, datetime
import sys

from rich import print

from toolbox import *
from mailtool import *


DESC = """
Access multiple IMAP mail accounts for showing filtered views of email and downloading
email attachments.
"""

EPILOG = """
A YAML file describing email account names, URLs and credentials must be available
at the path specified by the environment variable MAILTOOL_CFG_FILE.  The format
of the YAML file is:

account1name:
  url: mail.account1.com
  username: person@account1.com
  password: theP@ssw0rd

account2name:
  url: mail.account2.org
  username: otherperson@account2.org
  password: otherP@$$w0rD

etc.
"""


def main():
    parser = argparse.ArgumentParser(
        description=DESC, epilog=EPILOG, formatter_class=RawDescriptionHelpFormatter
    )
    parser.add_argument(
        "-a", "--account", default="all", help="Mail account name (or all)"
    )
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
    parser.add_argument(
        "-r",
        "--rules",
        action="store_true",
        default=False,
        help="Show automated account rules",
    )
    parser.add_argument(
        "-x",
        "--dryrun",
        action="store_true",
        default=False,
        help="Don't perform automated rules, just show rule processing as a dry run",
    )

    args = parser.parse_args()
    argsd = vars(args)
    if len(sys.argv) == 1:
        parser.print_usage()
        exit()

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
        print("Processing mail account: [#80FFC0 bold]%s[/]" % (account))
        if argsd["folders"]:
            m = MailAccount(account_name=account)
            for folder in m.folders:
                print(
                    "Mailbox : [#80FFC0 bold]%s[/] [white]/[/] [#20C040]%s[/]"
                    % (account, folder)
                )
        if argsd["rules"]:
            rules = MailSession.get_account_rules(account)
            if rules is not None:
                m = MailAccount(account)
                m.process_rules(dryrun=argsd["dryrun"])
            else:
                print("No automated rules are defined")

        if search_dict is not None and not argsd["folders"] and not argsd["rules"]:
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
                uids = m.get_message_uids(folder, search_dict)
                if argsd["unique"]:
                    m.unique_senders(uids, folder)
                elif argsd["output"]:
                    m.download_attachments(uids, folder)
                else:
                    objs = m.get_message_objs(uids, folder, show_attach)
                    m.print_messages(
                        objs,
                        show_attachments=show_attach,
                        show_name=argsd["name"],
                    )


if __name__ == "__main__":
    main()
