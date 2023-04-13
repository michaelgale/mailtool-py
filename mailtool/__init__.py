"""mailtool-py - A python package to access multiple IMAP email accounts."""

import os
import metayaml

# fmt: off
__project__ = 'mailtool'
__version__ = '0.1.1'
# fmt: on

VERSION = __project__ + "-" + __version__

script_dir = os.path.dirname(__file__)

FETCH_KEYS = [b"UID", b"ENVELOPE", b"BODYSTRUCTURE", b"FLAGS"]

from .helpers import (
    colour_address_str,
    colour_date_str,
    colour_file_str,
    colour_size_str,
    colour_subject_str,
    UID_COLOUR,
    DATE_COLOUR,
    FILE_COLOUR,
    FOLDER_COLOUR,
)
from .exceptions import *
from .mailattachment import MailAttachment
from .mailmessage import MailMessage
from .mailsession import MailSession
from .mailaccount import MailAccount
