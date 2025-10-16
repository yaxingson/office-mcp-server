import logging
from mcp.server import FastMCP

logging.basicConfig(
  level=logging.INFO
)

logger = logging.getLogger('office-mcp')

mcp = FastMCP(
  name='office-mcp-server'
)

