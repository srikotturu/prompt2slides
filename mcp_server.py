from mcp.server.fastmcp import FastMCP
import calendar_services

# Initialize FastMCP server
mcp = FastMCP("calendar")


@mcp.tool()
async def create_presentation():
    ...