import { OpenAIToolDefinition, ToolExecutionResult } from '../types';

export interface McpToolDiscoveryResult {
  serverUrl: string;
  tools: OpenAIToolDefinition[];
  error?: string;
}

export async function fetchMcpTools(serverUrl: string): Promise<McpToolDiscoveryResult> {
  const cleanUrl = serverUrl.trim().replace(/\/+$/, '');
  if (!cleanUrl) return { serverUrl, tools: [] };

  try {
    // Try standard JSON-RPC tools/list request
    const response = await fetch(`${cleanUrl}/tools/list`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/list',
        params: {},
      }),
    });

    if (!response.ok) {
      // Fallback: try GET /tools
      const fallbackResp = await fetch(`${cleanUrl}/tools`, { method: 'GET' });
      if (!fallbackResp.ok) {
        return { serverUrl, tools: [], error: `HTTP ${fallbackResp.status}: ${fallbackResp.statusText}` };
      }
      const data = await fallbackResp.json();
      const rawTools = Array.isArray(data.tools) ? data.tools : [];
      return {
        serverUrl,
        tools: rawTools.map((t: Record<string, unknown>) => formatMcpToolToOpenAI(cleanUrl, t)),
      };
    }

    const data = await response.json();
    const rawTools = data.result?.tools || data.tools || [];
    return {
      serverUrl,
      tools: rawTools.map((t: Record<string, unknown>) => formatMcpToolToOpenAI(cleanUrl, t)),
    };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.warn(`Failed to fetch MCP tools from ${cleanUrl}:`, err);
    return { serverUrl, tools: [], error: msg };
  }
}

function formatMcpToolToOpenAI(serverUrl: string, mcpTool: Record<string, unknown>): OpenAIToolDefinition {
  const name = String(mcpTool.name || 'unnamed_tool');
  const description = String(mcpTool.description || '');
  const inputSchema = (mcpTool.inputSchema as Record<string, unknown>) || { type: 'object', properties: {} };

  return {
    type: 'function',
    function: {
      name: `mcp_${name}`,
      description: `[MCP Tool from ${serverUrl}] ${description}`,
      parameters: inputSchema,
    },
    serverUrl,
  };
}

export async function executeMcpTool(
  serverUrl: string,
  toolName: string,
  args: Record<string, unknown>
): Promise<ToolExecutionResult> {
  const cleanUrl = serverUrl.trim().replace(/\/+$/, '');
  // Strip 'mcp_' prefix if added
  const rawToolName = toolName.startsWith('mcp_') ? toolName.slice(4) : toolName;

  try {
    const response = await fetch(`${cleanUrl}/tools/call`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        jsonrpc: '2.0',
        id: Date.now(),
        method: 'tools/call',
        params: {
          name: rawToolName,
          arguments: args,
        },
      }),
    });

    if (!response.ok) {
      return {
        success: false,
        message: `MCP Server Error (${response.status}): ${response.statusText}`,
      };
    }

    const data = await response.json();
    const result = data.result || data;

    return {
      success: true,
      message: `MCP Tool "${rawToolName}" executed successfully.`,
      data: result,
    };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return {
      success: false,
      message: `MCP Tool execution failed: ${msg}`,
    };
  }
}
