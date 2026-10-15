"""Mock of the Microsoft Graph v1.0 Outlook mail endpoints used by the pipeline.

Mirrors the real wire format: the same paths (under /v1.0), OData query
options ($top, $skip, $orderby, $select), the @odata.context / value envelope,
@odata.nextLink paging, the Prefer: outlook.body-content-type header and the
Graph error format.
"""

import html
import uuid
from datetime import datetime, timezone

from fastapi import Body, FastAPI, Header, Query, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse

from .mailbox import MESSAGE_FIELDS, Mailbox

app = FastAPI(title="Outlook (Microsoft Graph) mock")
mailbox = Mailbox()

SUPPORTED_QUERY_OPTIONS = {"$top", "$skip", "$orderby", "$select"}
SORTABLE_FIELDS = {
    "receivedDateTime", "sentDateTime", "createdDateTime", "lastModifiedDateTime",
    "subject", "isRead", "hasAttachments",
}


class GraphError(Exception):
    def __init__(self, status_code: int, code: str, message: str):
        self.status_code = status_code
        self.code = code
        self.message = message


@app.exception_handler(GraphError)
async def graph_error_handler(request: Request, exc: GraphError):
    return JSONResponse(status_code=exc.status_code, content={
        "error": {
            "code": exc.code,
            "message": exc.message,
            "innerError": {
                "date": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S"),
                "request-id": str(uuid.uuid4()),
            },
        }
    })


@app.exception_handler(RequestValidationError)
async def validation_error_handler(request: Request, exc: RequestValidationError):
    details = "; ".join(f"{'.'.join(map(str, e['loc']))}: {e['msg']}" for e in exc.errors())
    return await graph_error_handler(request, GraphError(status_code=400, code="BadRequest", message=details))


def item_not_found() -> GraphError:
    return GraphError(404, "ErrorItemNotFound", "The specified object was not found in the store.")


def odata_context(request: Request, path: str) -> str:
    return f"{request.base_url}v1.0/$metadata#users('{mailbox.user_id}')/{path}"


def reject_unsupported_query_options(request: Request) -> None:
    for key in request.query_params:
        if key.startswith("$") and key not in SUPPORTED_QUERY_OPTIONS:
            raise GraphError(400, "BadRequest", f"Query option '{key}' is not supported by this mock.")


def unknown_property(name: str) -> GraphError:
    return GraphError(400, "BadRequest", f"Could not find a property named '{name}' on type 'microsoft.graph.message'.")


def parse_select(select: str | None) -> list[str] | None:
    if not select:
        return None
    fields = [field.strip() for field in select.split(",") if field.strip()]
    for field in fields:
        if field not in MESSAGE_FIELDS:
            raise unknown_property(field)
    return fields


def sort_messages(messages: list[dict], orderby: str | None) -> list[dict]:
    # Without $orderby Outlook returns the newest messages first
    if not orderby:
        return sorted(messages, key=lambda m: m["receivedDateTime"], reverse=True)

    clauses = []
    for clause in orderby.split(","):
        parts = clause.split()
        if not parts or len(parts) > 2 or (len(parts) == 2 and parts[1].lower() not in ("asc", "desc")):
            raise GraphError(400, "BadRequest", f"Invalid $orderby clause: '{clause.strip()}'.")
        if parts[0] not in SORTABLE_FIELDS:
            raise unknown_property(parts[0])
        clauses.append((parts[0], len(parts) == 2 and parts[1].lower() == "desc"))

    # Stable sort, applied from the least to the most significant clause
    result = list(messages)
    for field, descending in reversed(clauses):
        result.sort(key=lambda m: m[field], reverse=descending)
    return result


def wants_text_body(prefer: str | None) -> bool:
    return prefer is not None and 'outlook.body-content-type="text"' in prefer.replace("'", '"')


def as_html(text: str) -> str:
    content = html.escape(text).replace("\n", "<br>\r\n")
    return (
        "<html>\r\n<head>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">\r\n"
        f"</head>\r\n<body>\r\n{content}\r\n</body>\r\n</html>\r\n"
    )


def render_message(message: dict, text_body: bool, select: list[str] | None = None) -> dict:
    result = dict(message)
    if not text_body:
        result["body"] = {"contentType": "html", "content": as_html(message["body"]["content"])}
    if select is not None:
        # id (and the derived type annotation) are always returned, like in the real API
        keep = {"@odata.type", "id", *select}
        result = {key: value for key, value in result.items() if key in keep}
    return result


def json_response(content: dict, text_body: bool, status_code: int = 200) -> JSONResponse:
    headers = {"Preference-Applied": 'outlook.body-content-type="text"'} if text_body else None
    return JSONResponse(content=content, status_code=status_code, headers=headers)


def message_page(request: Request, messages: list[dict], context_path: str,
                 top: int, skip: int, orderby: str | None, select: str | None, prefer: str | None) -> JSONResponse:
    reject_unsupported_query_options(request)
    fields = parse_select(select)
    text_body = wants_text_body(prefer)

    ordered = sort_messages(messages, orderby)
    page = ordered[skip:skip + top]

    if fields:
        context_path += f"({','.join(fields)})"
    content = {
        "@odata.context": odata_context(request, context_path),
        "value": [render_message(m, text_body, fields) for m in page],
    }
    if skip + top < len(ordered):
        content["@odata.nextLink"] = str(request.url.include_query_params(**{"$skip": skip + top}))
    return json_response(content, text_body)


# --- Graph endpoints ---------------------------------------------------------

TopParam = Query(10, alias="$top", ge=1, le=1000, description="Page size (Graph default: 10)")
SkipParam = Query(0, alias="$skip", ge=0)
OrderByParam = Query(None, alias="$orderby", examples=["receivedDateTime desc"])
SelectParam = Query(None, alias="$select", examples=["subject,from,receivedDateTime"])
PreferHeader = Header(None, description='e.g. outlook.body-content-type="text"')


@app.get("/v1.0/me")
def get_me(request: Request):
    return {
        "@odata.context": f"{request.base_url}v1.0/$metadata#users/$entity",
        **mailbox.user_profile(),
    }


@app.get("/v1.0/me/mailFolders")
def list_mail_folders(request: Request):
    return {
        "@odata.context": odata_context(request, "mailFolders"),
        "value": mailbox.list_folders(),
    }


@app.get("/v1.0/me/messages")
def list_messages(request: Request, top: int = TopParam, skip: int = SkipParam,
                  orderby: str | None = OrderByParam, select: str | None = SelectParam,
                  prefer: str | None = PreferHeader):
    return message_page(request, mailbox.list_messages(), "messages", top, skip, orderby, select, prefer)


@app.get("/v1.0/me/mailFolders/{folder_id}/messages")
def list_folder_messages(request: Request, folder_id: str, top: int = TopParam, skip: int = SkipParam,
                         orderby: str | None = OrderByParam, select: str | None = SelectParam,
                         prefer: str | None = PreferHeader):
    resolved = mailbox.resolve_folder_id(folder_id)
    if resolved is None:
        raise item_not_found()
    return message_page(request, mailbox.list_messages(resolved), f"mailFolders('{folder_id}')/messages",
                        top, skip, orderby, select, prefer)


@app.get("/v1.0/me/messages/{message_id}")
def get_message(request: Request, message_id: str, select: str | None = SelectParam,
                prefer: str | None = PreferHeader):
    reject_unsupported_query_options(request)
    message = mailbox.get_message(message_id)
    if message is None:
        raise item_not_found()
    fields = parse_select(select)
    text_body = wants_text_body(prefer)
    context_path = f"messages({','.join(fields)})/$entity" if fields else "messages/$entity"
    content = {"@odata.context": odata_context(request, context_path), **render_message(message, text_body, fields)}
    return json_response(content, text_body)


@app.post("/v1.0/me/messages/{message_id}/move", status_code=201)
def move_message(request: Request, message_id: str,
                 destination_id: str = Body(..., alias="destinationId", embed=True,
                                            description="Folder id or well-known folder name"),
                 prefer: str | None = PreferHeader):
    if mailbox.get_message(message_id) is None:
        raise item_not_found()
    destination = mailbox.resolve_folder_id(destination_id)
    if destination is None:
        raise item_not_found()

    moved = mailbox.move_message(message_id, destination)
    text_body = wants_text_body(prefer)
    content = {
        "@odata.context": f"{request.base_url}v1.0/$metadata#message",
        "@odata.type": "#microsoft.graph.message",
        **render_message(moved, text_body),
    }
    return json_response(content, text_body, status_code=201)


# --- Mock-only helpers (not part of Graph) -----------------------------------

@app.post("/mock/reset", tags=["mock"])
def reset_mailbox():
    """Reload the fixture: every message back in its original folder, with fresh ids and dates."""
    mailbox.reset()
    return {"status": "reset", "messages": len(mailbox.messages)}
