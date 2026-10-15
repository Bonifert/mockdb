"""In-memory mailbox state built from the fixture file.

The fixture (data/mailbox.json) only holds static content. Everything that the
real Graph API generates or changes at runtime is derived here:

- relative dates ("now-6h", "now+4d@09:00") are resolved to UTC timestamps
- message ids are generated, and regenerated on every move (like the real API)
- sender, sentDateTime, createdDateTime, lastModifiedDateTime, bodyPreview,
  parentFolderId, isDraft and bccRecipients are filled in
"""

import base64
import copy
import hashlib
import json
import re
import secrets
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path

FIXTURE_PATH = Path(__file__).resolve().parent.parent / "data" / "mailbox.json"

# Well-known folder names that can be used instead of a folder id, e.g. /me/mailFolders/inbox
WELL_KNOWN_FOLDERS = {"inbox"}

RELATIVE_DATE = re.compile(r"^now(?P<offset>[+-]\d+)(?P<unit>[mhd])(?:@(?P<hour>\d{2}):(?P<minute>\d{2}))?$")
UNITS = {"m": "minutes", "h": "hours", "d": "days"}

# Field order of a real Graph message response
MESSAGE_FIELDS = [
    "id", "createdDateTime", "lastModifiedDateTime", "categories", "receivedDateTime", "sentDateTime",
    "hasAttachments", "subject", "bodyPreview", "importance", "parentFolderId", "conversationId",
    "isRead", "isDraft", "body", "sender", "from", "toRecipients", "ccRecipients", "bccRecipients",
    "flag", "meetingMessageType", "meetingInfo",
]


def resolve_relative_date(value: str, now: datetime) -> datetime:
    match = RELATIVE_DATE.match(value)
    if not match:
        raise ValueError(f"Invalid relative date: {value!r}")
    result = now + timedelta(**{UNITS[match["unit"]]: int(match["offset"])})
    if match["hour"] is not None:
        result = result.replace(hour=int(match["hour"]), minute=int(match["minute"]), second=0)
    return result


def format_timestamp(value: datetime) -> str:
    """Message timestamps, e.g. 2026-10-01T08:15:00Z"""
    return value.strftime("%Y-%m-%dT%H:%M:%SZ")


def format_event_datetime(value: datetime) -> str:
    """dateTimeTimeZone.dateTime of events, e.g. 2026-10-05T09:00:00.0000000"""
    return value.strftime("%Y-%m-%dT%H:%M:%S.0000000")


def new_message_id() -> str:
    return "AAMkAD" + base64.urlsafe_b64encode(secrets.token_bytes(96)).decode()


def folder_id(key: str) -> str:
    # Stable across restarts, unlike message ids
    return "AQMkAD" + base64.urlsafe_b64encode(hashlib.sha256(key.encode()).digest()[:24]).decode()


class Mailbox:
    def __init__(self, fixture_path: Path = FIXTURE_PATH):
        self.fixture_path = fixture_path
        self.reset()

    def reset(self) -> None:
        """(Re)load the fixture, so a pipeline run can start from a clean state."""
        fixture = json.loads(self.fixture_path.read_text(encoding="utf-8"))
        now = datetime.now(timezone.utc).replace(microsecond=0)

        self.owner = fixture["owner"]
        self.user_id = str(uuid.uuid5(uuid.NAMESPACE_DNS, self.owner["address"]))
        self.root_folder_id = folder_id("msgfolderroot")

        self.folders = {}
        self.folder_keys = {}
        for folder in fixture["folders"]:
            fid = folder_id(folder["key"])
            self.folder_keys[folder["key"]] = fid
            self.folders[fid] = {
                "id": fid,
                "displayName": folder["displayName"],
                "parentFolderId": self.root_folder_id,
                "childFolderCount": 0,
                "isHidden": False,
            }

        self.messages = {}
        for raw in fixture["messages"]:
            message = self._build_message(raw, now)
            self.messages[message["id"]] = message

    def _build_message(self, raw: dict, now: datetime) -> dict:
        raw = copy.deepcopy(raw)
        received = resolve_relative_date(raw.pop("receivedDateTime"), now)
        folder_key = raw.pop("folder")

        message = {
            **raw,
            "id": new_message_id(),
            "createdDateTime": format_timestamp(received),
            "lastModifiedDateTime": format_timestamp(received),
            "receivedDateTime": format_timestamp(received),
            "sentDateTime": format_timestamp(received - timedelta(seconds=2)),
            "bodyPreview": raw["body"]["content"][:255],
            "parentFolderId": self.folder_keys[folder_key],
            "isDraft": False,
            "sender": raw["from"],
            "bccRecipients": [],
        }

        if "meetingInfo" in message:
            for field in ("start", "end"):
                value = message["meetingInfo"][field]
                value["dateTime"] = format_event_datetime(resolve_relative_date(value["dateTime"], now))

        ordered = {field: message[field] for field in MESSAGE_FIELDS if field in message}
        if "meetingMessageType" in message:
            ordered = {"@odata.type": "#microsoft.graph.eventMessageRequest", **ordered}
        return ordered

    # --- user ----------------------------------------------------------------

    def user_profile(self) -> dict:
        """The mailbox owner, shaped like the default GET /me response."""
        given_name, _, surname = self.owner["name"].partition(" ")
        return {
            "businessPhones": [],
            "displayName": self.owner["name"],
            "givenName": given_name,
            "jobTitle": None,
            "mail": self.owner["address"],
            "mobilePhone": None,
            "officeLocation": None,
            "preferredLanguage": None,
            "surname": surname,
            "userPrincipalName": self.owner["address"],
            "id": self.user_id,
        }

    # --- folders -------------------------------------------------------------

    def resolve_folder_id(self, id_or_name: str) -> str | None:
        """Accepts a folder id or a well-known folder name."""
        if id_or_name in self.folders:
            return id_or_name
        if id_or_name.lower() in WELL_KNOWN_FOLDERS:
            return self.folder_keys.get(id_or_name.lower())
        return None

    def list_folders(self) -> list[dict]:
        result = []
        for folder in self.folders.values():
            in_folder = [m for m in self.messages.values() if m["parentFolderId"] == folder["id"]]
            result.append({
                **folder,
                "unreadItemCount": sum(1 for m in in_folder if not m["isRead"]),
                "totalItemCount": len(in_folder),
            })
        return result

    # --- messages ------------------------------------------------------------

    def list_messages(self, folder_id: str | None = None) -> list[dict]:
        return [
            m for m in self.messages.values()
            if folder_id is None or m["parentFolderId"] == folder_id
        ]

    def get_message(self, message_id: str) -> dict | None:
        return self.messages.get(message_id)

    def move_message(self, message_id: str, destination_folder_id: str) -> dict:
        """Like the real API: the message gets a new id, the old id stops working."""
        message = self.messages.pop(message_id)
        moved = {
            **message,
            "id": new_message_id(),
            "parentFolderId": destination_folder_id,
            "lastModifiedDateTime": format_timestamp(datetime.now(timezone.utc)),
        }
        self.messages[moved["id"]] = moved
        return moved
