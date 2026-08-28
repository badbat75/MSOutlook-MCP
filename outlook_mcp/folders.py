"""Resolving and rendering Outlook mail folders.

Graph accepts a handful of well-known folder aliases directly, but a user says
"Centri Estivi", not an opaque 120-character folder id. Everything that turns
one into the other lives here.
"""

from .auth import GraphClient

# Well-known folder aliases Graph accepts directly (no ID lookup needed).
WELL_KNOWN_FOLDERS = {
    "inbox": "inbox",
    "archive": "archive",
    "deleteditems": "deletedItems",
    "trash": "deletedItems",
    "junkemail": "junkEmail",
    "junk": "junkEmail",
    "drafts": "drafts",
    "sentitems": "sentItems",
    "sent": "sentItems",
}


async def find_folder_id_by_name(
    graph: GraphClient, name: str, parent: str = "/me/mailFolders"
) -> str:
    """Depth-first search of the folder tree for a folder matching displayName.

    Returns the folder ID, or None if no match is found. Searches nested
    child folders too, so subfolders (e.g. a folder inside Inbox) are found.
    """
    data = await graph.get(
        parent,
        params={"$top": 100, "$select": "id,displayName,childFolderCount"},
    )
    folders = data.get("value", [])
    target = name.strip().lower()
    for f in folders:
        if f.get("displayName", "").lower() == target:
            return f["id"]
    for f in folders:
        if f.get("childFolderCount", 0) > 0:
            found = await find_folder_id_by_name(
                graph, name, f"/me/mailFolders/{f['id']}/childFolders"
            )
            if found:
                return found
    return None


async def resolve_folder(graph: GraphClient, name: str) -> str:
    """Resolve a folder reference to a value Graph accepts (well-known name or ID).

    Accepts well-known aliases (inbox, archive, ...), a raw folder ID, or a
    user-facing display name (including nested subfolders). Raises ValueError
    with a helpful message if a display name cannot be matched.
    """
    key = name.strip().lower()
    if key in WELL_KNOWN_FOLDERS:
        return WELL_KNOWN_FOLDERS[key]
    # Folder IDs are long opaque strings; short values are treated as names.
    if len(name) > 60 and " " not in name:
        return name
    folder_id = await find_folder_id_by_name(graph, name)
    if folder_id:
        return folder_id
    raise ValueError(
        f"Folder '{name}' not found. Use outlook_list_folders to see available "
        f"folders (names are case-insensitive; subfolders are supported)."
    )


async def format_folder_tree(
    graph: GraphClient, parent: str, top: int, recurse: bool, depth: int = 0
) -> str:
    """Render mail folders under `parent` as a nested bullet list.

    When `recurse` is True, folders with children are expanded depth-first so
    subfolders (e.g. a folder under Inbox) appear indented beneath their parent.
    """
    data = await graph.get(
        parent,
        params={"$top": top, "$select": "id,displayName,totalItemCount,unreadItemCount,childFolderCount"},
    )
    folders = data.get("value", [])
    lines = ""
    indent = "  " * depth
    for f in folders:
        unread = f.get("unreadItemCount", 0)
        unread_badge = f" (📬 {unread} unread)" if unread > 0 else ""
        lines += (
            f"{indent}- **{f['displayName']}**{unread_badge}: "
            f"{f.get('totalItemCount', 0)} items | ID: `{f['id']}`\n"
        )
        if recurse and f.get("childFolderCount", 0) > 0:
            lines += await format_folder_tree(
                graph, f"/me/mailFolders/{f['id']}/childFolders", top, recurse, depth + 1
            )
    return lines
