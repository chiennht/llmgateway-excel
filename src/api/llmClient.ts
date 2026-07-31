import { AppSettings, ChatMessage, OpenAIToolCall } from '../types';
import { EXCEL_TOOLS_SCHEMA, executeExcelTool } from '../excel/tools';

export interface LLMStreamOptions {
  settings: AppSettings;
  messages: ChatMessage[];
  onMessageUpdate: (messages: ChatMessage[]) => void;
  onStatusChange?: (status: string) => void;
}

export async function runLLMChatLoop(options: LLMStreamOptions): Promise<ChatMessage[]> {
  const { settings, messages, onMessageUpdate, onStatusChange } = options;

  let cleanBaseUrl = settings.baseUrl.trim().replace(/\/+$/, '');
  if (!cleanBaseUrl.endsWith('/chat/completions')) {
    if (!cleanBaseUrl.endsWith('/v1')) {
      cleanBaseUrl += '/v1';
    }
    cleanBaseUrl += '/chat/completions';
  }

  const history: ChatMessage[] = [...messages];
  const maxToolLoops = 10;
  let loopCount = 0;

  while (loopCount < maxToolLoops) {
    loopCount++;
    if (onStatusChange) {
      onStatusChange(loopCount > 1 ? `Running tool step ${loopCount}...` : 'Generating response...');
    }

    const payloadMessages = history.map((msg) => {
      const pMsg: Record<string, unknown> = {
        role: msg.role,
        content: msg.content ?? '',
      };
      if (msg.name) pMsg.name = msg.name;
      if (msg.tool_call_id) pMsg.tool_call_id = msg.tool_call_id;
      if (msg.tool_calls) pMsg.tool_calls = msg.tool_calls;
      return pMsg;
    });

    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
    };
    if (settings.apiKey.trim()) {
      headers['Authorization'] = `Bearer ${settings.apiKey.trim()}`;
    }

    const requestBody = {
      model: settings.model.trim() || 'gpt-4o',
      messages: payloadMessages,
      tools: EXCEL_TOOLS_SCHEMA,
      tool_choice: 'auto',
      stream: false,
    };

    let response: Response;
    try {
      response = await fetch(cleanBaseUrl, {
        method: 'POST',
        headers,
        body: JSON.stringify(requestBody),
      });
    } catch (fetchErr: unknown) {
      const errText = fetchErr instanceof Error ? fetchErr.message : String(fetchErr);
      const errMsg: ChatMessage = {
        id: `err-${Date.now()}`,
        role: 'assistant',
        content: `**Network / API Error:** Failed to connect to \`${cleanBaseUrl}\`.\n\nDetail: ${errText}`,
        timestamp: Date.now(),
      };
      history.push(errMsg);
      onMessageUpdate(history);
      throw fetchErr;
    }

    if (!response.ok) {
      const errorText = await response.text();
      const errMsg: ChatMessage = {
        id: `err-${Date.now()}`,
        role: 'assistant',
        content: `**API Error (${response.status}):** ${response.statusText}\n\`\`\`json\n${errorText}\n\`\`\``,
        timestamp: Date.now(),
      };
      history.push(errMsg);
      onMessageUpdate(history);
      return history;
    }

    const json = await response.json();
    const choice = json.choices?.[0];
    if (!choice || !choice.message) {
      const errMsg: ChatMessage = {
        id: `err-${Date.now()}`,
        role: 'assistant',
        content: '**Invalid API Response:** Model returned empty response or invalid payload format.',
        timestamp: Date.now(),
      };
      history.push(errMsg);
      onMessageUpdate(history);
      return history;
    }

    const assistantMsgObj = choice.message;
    const reasoningText = assistantMsgObj.reasoning_content || assistantMsgObj.reasoning || undefined;
    const textContent = assistantMsgObj.content || null;
    const toolCalls: OpenAIToolCall[] | undefined = assistantMsgObj.tool_calls;

    const assistantMessage: ChatMessage = {
      id: `asst-${Date.now()}-${loopCount}`,
      role: 'assistant',
      content: textContent,
      reasoning: reasoningText,
      tool_calls: toolCalls,
      timestamp: Date.now(),
    };

    history.push(assistantMessage);
    onMessageUpdate(history);

    if (!toolCalls || toolCalls.length === 0) {
      // Done! Final response received without tool calls
      if (onStatusChange) onStatusChange('Ready');
      return history;
    }

    // Process tool calls
    for (const toolCall of toolCalls) {
      const fnName = toolCall.function.name;
      let fnArgs: Record<string, unknown> = {};
      try {
        fnArgs = JSON.parse(toolCall.function.arguments || '{}');
      } catch (e) {
        console.warn(`Failed to parse tool arguments for ${fnName}:`, e);
      }

      if (onStatusChange) {
        onStatusChange(`Executing Excel tool: ${fnName}...`);
      }

      const toolResult = await executeExcelTool(fnName, fnArgs);

      const toolMessage: ChatMessage = {
        id: `tool-${Date.now()}-${toolCall.id}`,
        role: 'tool',
        name: fnName,
        tool_call_id: toolCall.id,
        content: JSON.stringify(toolResult),
        timestamp: Date.now(),
      };

      history.push(toolMessage);
      onMessageUpdate(history);
    }
  }

  if (onStatusChange) onStatusChange('Ready');
  return history;
}
