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
# MailSession class - acts as a context manager to access an IMAP server

import imapclient

from mailtool import *

try:
    MAILTOOL_CFG_FILE = os.path.expanduser(os.environ["MAILTOOL_CFG_FILE"])
    MAILTOOL_ACCOUNTS = metayaml.read(MAILTOOL_CFG_FILE, disable_order_dict=True)
except:
    print("mailtool configuration file needs to be specified")
    print("with the environment variable:  MAILTOOL_CFG_FILE")


class MailSession:
    def __init__(self, account_name=None):
        self.session = None
        if account_name is not None:
            self.account_name = account_name.lower()
        else:
            raise MissingAccountNameException()
        if account_name in MAILTOOL_ACCOUNTS:
            account = MAILTOOL_ACCOUNTS[account_name]
            self.account_url = account["url"]
            self.username = account["username"]
            self.pwd = account["password"]
            return
        raise UnknownAccountNameException()

    def __enter__(self):
        try:
            self.session = imapclient.IMAPClient(
                self.account_url, use_uid=True, ssl=True
            )
            self.session.login(self.username, self.pwd)
        except:
            raise AccountLoginFailedException()
        return self.session

    def __exit__(self, exc_type, exc_value, exc_traceback):
        if self.session is not None:
            self.session.logout()

    @staticmethod
    def get_accounts():
        return [k for k in MAILTOOL_ACCOUNTS.keys()]

    @staticmethod
    def get_account_rules(account):
        if not account in MAILTOOL_ACCOUNTS:
            raise UnknownAccountNameException
        a = MAILTOOL_ACCOUNTS[account]
        if "rules" in a:
            rules = a["rules"]
            if not isinstance(rules, list):
                rules = [rules]
            return [os.path.expanduser(f) for f in rules]
        return None
