"""Walks through every endpoint of the running mock and prints what the data looks like.

Start the server first (python -m mock_outlook), then run:  python demo.py [base_url]
The mailbox is reset at the start and at the end, so the run doesn't leave anything behind.
"""

import json
import sys
import urllib.error
import urllib.parse
import urllib.request

# 127.0.0.1 instead of localhost: on Windows "localhost" tries IPv6 first and waits ~2s per request
BASE = (sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8000").rstrip("/")
GRAPH = BASE + "/v1.0"
TEXT_BODY = {"Prefer": 'outlook.body-content-type="text"'}

sys.stdout.reconfigure(encoding="utf-8")


# --- HTTP ----------------------------------------------------------------------

def call(method: str, url: str, params: dict | None = None, body: dict | None = None,
         headers: dict | None = None) -> tuple[int, dict]:
    if not url.startswith("http"):
        url = GRAPH + url
    if params:
        url += "?" + urllib.parse.urlencode(params, quote_via=urllib.parse.quote)

    print(f"\n  → {method} {urllib.parse.unquote(url)}")
    if headers:
        for name, value in headers.items():
            print(f"    {name}: {value}")
    if body is not None:
        print(f"    body: {json.dumps(body)}")

    request = urllib.request.Request(url, method=method, data=json.dumps(body).encode() if body is not None else None,
                                     headers={"Content-Type": "application/json", **(headers or {})})
    try:
        with urllib.request.urlopen(request) as response:
            status, data = response.status, json.load(response)
    except urllib.error.HTTPError as error:
        status, data = error.code, json.load(error)
    except urllib.error.URLError as error:
        sys.exit(f"\nCan't reach the mock at {BASE} ({error.reason}). Start it with: python -m mock_outlook")

    print(f"  ← {status}")
    return status, data


# --- printing --------------------------------------------------------------------

def section(title: str) -> None:
    print("\n\n" + "=" * 100)
    print(f"  {title}")
    print("=" * 100)


def show_json(data: dict, max_string: int = 120) -> None:
    """Pretty-print JSON, shortening long strings (ids, bodies) so the structure stays readable."""
    def shorten(value):
        if isinstance(value, dict):
            return {k: shorten(v) for k, v in value.items()}
        if isinstance(value, list):
            return [shorten(v) for v in value]
        if isinstance(value, str) and len(value) > max_string:
            return value[:max_string] + f"… (+{len(value) - max_string} chars)"
        return value
    print("\n" + "\n".join("    " + line for line in json.dumps(shorten(data), indent=2, ensure_ascii=False).splitlines()))


def short_id(value: str) -> str:
    return value[:14] + "…"


def table(rows: list[list], headers: list[str]) -> None:
    widths = [max(len(str(x)) for x in column) for column in zip(headers, *rows)]
    line = lambda cells: "    " + "  ".join(str(c).ljust(w) for c, w in zip(cells, widths))
    print("\n" + line(headers))
    print("    " + "  ".join("-" * w for w in widths))
    for row in rows:
        print(line(row))


def message_rows(messages: list[dict], folder_names: dict) -> list[list]:
    return [[
        i,
        m["receivedDateTime"],
        "meeting" if "meetingMessageType" in m else "mail",
        "" if m["isRead"] else "unread",
        m["from"]["emailAddress"]["name"],
        m["subject"][:50],
        folder_names.get(m["parentFolderId"], "?"),
    ] for i, m in enumerate(messages, 1)]


MESSAGE_HEADERS = ["#", "received (UTC)", "type", "read", "from", "subject", "folder"]
LIST_SELECT = "subject,from,receivedDateTime,isRead,parentFolderId,meetingMessageType"


def show_folders() -> dict:
    _, data = call("GET", "/me/mailFolders")
    table([[f["displayName"], f["totalItemCount"], f["unreadItemCount"], short_id(f["id"])] for f in data["value"]],
          ["folder", "total", "unread", "id"])
    return data


# --- walkthrough ---------------------------------------------------------------------

section("0. Reset the mock mailbox (mock-only endpoint)")
_, data = call("POST", BASE + "/mock/reset")
show_json(data)

section("1. GET /me: the signed-in user (the mailbox owner)")
_, me = call("GET", "/me")
show_json(me)

section("2. GET /me/mailFolders: the folders")
folders = show_folders()
print("\n  Raw response:")
show_json(folders)
folder_ids = {f["displayName"]: f["id"] for f in folders["value"]}
folder_names = {fid: name for name, fid in folder_ids.items()}

section("3. GET /me/mailFolders/inbox/messages: the latest 15 in the Inbox")
params = {"$orderby": "receivedDateTime desc", "$top": 15, "$select": LIST_SELECT}
_, page = call("GET", "/me/mailFolders/inbox/messages", params=params)
latest = page["value"]
table(message_rows(latest, folder_names), MESSAGE_HEADERS)
print(f"\n  Envelope keys: {list(page)}")
print(f"  @odata.context:  {page['@odata.context']}")
print(f"  @odata.nextLink: {urllib.parse.unquote(page.get('@odata.nextLink', '-'))}")
print("\n  First item as returned ($select keeps only id + the selected fields):")
show_json(latest[0])

section("4. Follow @odata.nextLink: the next page (older Inbox messages)")
_, next_page = call("GET", page["@odata.nextLink"])
table(message_rows(next_page["value"], folder_names), MESSAGE_HEADERS)
print(f"\n  @odata.nextLink: {next_page.get('@odata.nextLink', '- (no more pages)')}")

section("5. GET /me/messages: whole mailbox, default paging (10 per page)")
_, everything = call("GET", "/me/messages", params={"$select": LIST_SELECT})
table(message_rows(everything["value"], folder_names), MESSAGE_HEADERS)
print(f"\n  {len(everything['value'])} items, @odata.nextLink: {urllib.parse.unquote(everything.get('@odata.nextLink', '-'))}")

meeting = next(m for m in latest if "meetingMessageType" in m)
mail = next(m for m in latest if "meetingMessageType" not in m)

section(f"6. GET /me/messages/{{id}}: a full meeting request ('{meeting['subject']}')")
_, full_meeting = call("GET", f"/me/messages/{meeting['id']}", headers=TEXT_BODY)
show_json(full_meeting)

section(f"7. GET /me/messages/{{id}}: a full plain message ('{mail['subject']}')")
_, full_mail = call("GET", f"/me/messages/{mail['id']}", headers=TEXT_BODY)
show_json(full_mail)
print("\n  The complete body (text):\n")
print("\n".join("    | " + line for line in full_mail["body"]["content"].splitlines()))

section("8. The same message without the Prefer header: body comes back as HTML (Graph default)")
_, html_mail = call("GET", f"/me/messages/{mail['id']}", params={"$select": "subject,body"})
show_json(html_mail, max_string=400)

section("9. POST /me/messages/{id}/move: Inbox → Processing → Done")
_, moved = call("POST", f"/me/messages/{mail['id']}/move", body={"destinationId": folder_ids["Processing"]})
print(f"\n  old id:         {short_id(mail['id'])}")
print(f"  new id:         {short_id(moved['id'])}   ← the id changes on every move")
print(f"  parentFolderId: {folder_names[moved['parentFolderId']]}")
print(f"  lastModified:   {moved['lastModifiedDateTime']}")
print("\n  Raw response:")
show_json(moved, max_string=60)

print("\n  The old id no longer works:")
_, error = call("GET", f"/me/messages/{mail['id']}")
show_json(error)

print("\n  Folder counts after the move:")
show_folders()

_, done = call("POST", f"/me/messages/{moved['id']}/move", body={"destinationId": folder_ids["Done"]})
print(f"\n  Moved on to: {folder_names[done['parentFolderId']]} (id changed again: {short_id(done['id'])})")
show_folders()

section("10. Errors come back in the Graph error format")
for description, params in [
    ("unsupported query option", {"$filter": "isRead eq false"}),
    ("unknown field in $select", {"$select": "foo"}),
    ("$top out of range", {"$top": 0}),
]:
    print(f"\n  {description}:")
    _, error = call("GET", "/me/messages", params=params)
    show_json(error)

section("11. Reset, so the mailbox is back in its original state")
_, data = call("POST", BASE + "/mock/reset")
show_json(data)
show_folders()
print()
