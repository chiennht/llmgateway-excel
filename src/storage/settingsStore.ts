import { AppSettings } from '../types';

const STORAGE_KEY = 'llmgateway_excel_settings';

export const DEFAULT_SETTINGS: AppSettings = {
  baseUrl: 'https://api.openai.com/v1',
  apiKey: '',
  model: 'gpt-4o',
  systemPrompt: `You are an expert Excel AI assistant powered by an LLM Gateway.
You help users analyze spreadsheets, write formulas, clean data, and format workbooks.

CRITICAL DIRECTIVES:
1. ALWAYS use native function calling (tools) to execute actions on the Excel workbook (write_cells, read_range, format_range, clear_range, get_workbook_overview).
2. DO NOT write JSON blocks or text representations of tool calls in your message body. Execute tool calls directly via function calling!
3. Always inspect cell range data using read_range or get_workbook_overview before assuming cell locations.
4. Output clear explanations and summary tables using Markdown after performing spreadsheet actions.
5. If data or cell values are missing or ambiguous, report UNKNOWN or ask for clarification.`,
  mcpServers: [],
  activeSkillIds: ['data_cleaning', 'financial_analysis', 'formula_expert'],
};

export function loadSettings(): AppSettings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return DEFAULT_SETTINGS;
    const parsed = JSON.parse(raw);
    return {
      baseUrl: parsed.baseUrl || DEFAULT_SETTINGS.baseUrl,
      apiKey: parsed.apiKey || DEFAULT_SETTINGS.apiKey,
      model: parsed.model || DEFAULT_SETTINGS.model,
      systemPrompt: parsed.systemPrompt || DEFAULT_SETTINGS.systemPrompt,
      mcpServers: Array.isArray(parsed.mcpServers) ? parsed.mcpServers : DEFAULT_SETTINGS.mcpServers,
      activeSkillIds: Array.isArray(parsed.activeSkillIds) ? parsed.activeSkillIds : DEFAULT_SETTINGS.activeSkillIds,
    };
  } catch (err) {
    console.error('Failed to load settings from localStorage:', err);
    return DEFAULT_SETTINGS;
  }
}

export function saveSettings(settings: AppSettings): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
  } catch (err) {
    console.error('Failed to save settings to localStorage:', err);
  }
}
