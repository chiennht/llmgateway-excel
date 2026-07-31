import { AppSettings } from '../types';

const STORAGE_KEY = 'llmgateway_excel_settings';

export const DEFAULT_SETTINGS: AppSettings = {
  baseUrl: 'https://api.openai.com/v1',
  apiKey: '',
  model: 'gpt-4o',
  systemPrompt: `You are an expert Excel AI assistant powered by an LLM Gateway.
You help users analyze spreadsheets, write formulas, clean data, and format workbooks.

RULES & INSTRUCTIONS:
1. Always analyze context and requirements step-by-step.
2. Use tools to read or inspect range data before making assumptions.
3. When writing formulas or modifying cells, double check range coordinates and validity.
4. Output concise, helpful explanations with proper markdown.
5. If data or cell values are missing or unclear, report UNKNOWN or ask for clarification instead of inventing data.`,
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
