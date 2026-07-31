export interface AppSettings {
  baseUrl: string;
  apiKey: string;
  model: string;
  systemPrompt: string;
}

export interface ChatMessage {
  id: string;
  role: 'system' | 'user' | 'assistant' | 'tool';
  content: string | null;
  reasoning?: string;
  name?: string;
  tool_call_id?: string;
  tool_calls?: OpenAIToolCall[];
  timestamp: number;
}

export interface OpenAIToolCall {
  id: string;
  type: 'function';
  function: {
    name: string;
    arguments: string;
  };
}

export interface OpenAIToolDefinition {
  type: 'function';
  function: {
    name: string;
    description: string;
    parameters: Record<string, unknown>;
  };
}

export interface WorkbookOverview {
  activeSheet: string;
  sheets: string[];
  selectionAddress: string;
  tables?: string[];
}

export interface ToolExecutionResult {
  success: boolean;
  message: string;
  data?: unknown;
}
