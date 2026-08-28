"""Profile tool: who the server is authenticated as."""

from mcp.server.mcpserver import Context

from ..app import mcp
from ..credentials import get_graph
from ..helpers import handle_graph_error


@mcp.tool(
    name="outlook_get_profile",
    annotations={
        "title": "Get User Profile",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    },
)
async def outlook_get_profile(ctx: Context = None) -> str:
    """Get the authenticated user's profile information.

    Returns:
        str: User profile with name, email, job title, etc.
    """
    try:
        graph = get_graph(ctx)
        data = await graph.get(
            "/me",
            params={"$select": "displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone"},
        )
        result = "👤 **User Profile**\n\n"
        result += f"**Name:** {data.get('displayName', 'N/A')}\n"
        result += f"**Email:** {data.get('mail') or data.get('userPrincipalName', 'N/A')}\n"
        result += f"**Job Title:** {data.get('jobTitle', 'N/A')}\n"
        result += f"**Department:** {data.get('department', 'N/A')}\n"
        result += f"**Office:** {data.get('officeLocation', 'N/A')}\n"
        result += f"**Phone:** {data.get('mobilePhone', 'N/A')}\n"
        return result
    except Exception as e:
        return handle_graph_error(e)
