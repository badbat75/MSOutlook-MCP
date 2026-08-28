"""Attaching local files to a Graph message.

Graph takes an attachment inline up to 3MB; past that it wants an upload
session written in chunks. Which of the two applies is a size check, so
callers just hand over paths.
"""

import base64
import mimetypes
from pathlib import Path
from typing import List

import httpx

from .auth import GraphClient

# Graph accepts attachments up to 3MB inline on a message; larger files must go
# through an upload session uploaded in chunks (multiples of 320 KB recommended).
INLINE_ATTACHMENT_LIMIT = 3 * 1024 * 1024
UPLOAD_CHUNK_SIZE = 5 * 1024 * 1024


def read_attachment_meta(path: str):
    """Resolve a local file path to (Path, name, content_type, size)."""
    p = Path(path).expanduser()
    if not p.is_file():
        raise ValueError(f"Attachment file not found: {path}")
    content_type = mimetypes.guess_type(p.name)[0] or "application/octet-stream"
    return p, p.name, content_type, p.stat().st_size


async def attach_small_file(
    graph: GraphClient, message_id: str, p: Path, name: str, content_type: str
) -> None:
    """Attach a small (<=3MB) file inline via POST .../attachments."""
    content_b64 = base64.b64encode(p.read_bytes()).decode("ascii")
    await graph.post(
        f"/me/messages/{message_id}/attachments",
        json_data={
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": name,
            "contentType": content_type,
            "contentBytes": content_b64,
        },
    )


async def attach_large_file(
    graph: GraphClient, message_id: str, p: Path, name: str, content_type: str, size: int
) -> None:
    """Attach a large (>3MB) file via an upload session, uploaded in chunks."""
    session = await graph.post(
        f"/me/messages/{message_id}/attachments/createUploadSession",
        json_data={
            "AttachmentItem": {
                "attachmentType": "file",
                "name": name,
                "size": size,
                "contentType": content_type,
            }
        },
    )
    upload_url = session["uploadUrl"]
    # The upload URL is pre-authenticated; use a bare client (no Bearer header).
    async with httpx.AsyncClient(timeout=120.0) as client:
        with p.open("rb") as fh:
            start = 0
            while start < size:
                chunk = fh.read(UPLOAD_CHUNK_SIZE)
                end = start + len(chunk) - 1
                resp = await client.put(
                    upload_url,
                    content=chunk,
                    headers={
                        "Content-Length": str(len(chunk)),
                        "Content-Range": f"bytes {start}-{end}/{size}",
                        "Content-Type": "application/octet-stream",
                    },
                )
                resp.raise_for_status()
                start = end + 1


async def attach_files(graph: GraphClient, message_id: str, paths: List[str]) -> List[str]:
    """Attach each local file path to an existing message. Returns attached names."""
    names: List[str] = []
    for path in paths:
        p, name, content_type, size = read_attachment_meta(path)
        if size <= INLINE_ATTACHMENT_LIMIT:
            await attach_small_file(graph, message_id, p, name, content_type)
        else:
            await attach_large_file(graph, message_id, p, name, content_type, size)
        names.append(name)
    return names
