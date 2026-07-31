import { Skill } from '../types';

export const BUILTIN_SKILLS: Skill[] = [
  {
    id: 'data_cleaning',
    name: 'Data Cleaning & Normalization',
    category: 'Data Prep',
    description: 'Auto-detect messy entries, trim whitespace, fix dates, normalize emails & phones, and remove duplicates.',
    isBuiltIn: true,
    instructions: `[SPECIALIZED SKILL: Data Cleaning & Normalization]
- Always inspect table headers and sample rows first using read_range.
- Identify trailing/leading whitespace, mixed date formats (YYYY-MM-DD vs MM/DD/YYYY), or invalid text cases.
- Use write_cells to apply cleaned values directly to cells.
- Summarize changes made (e.g. number of trimmed strings, normalized dates).`,
  },
  {
    id: 'financial_analysis',
    name: 'Financial Analysis & Ratio Modeling',
    category: 'Finance',
    description: 'Calculate financial ratios (Profit Margin, ROI, CAGR), DCF valuation models, and income statement summaries.',
    isBuiltIn: true,
    instructions: `[SPECIALIZED SKILL: Financial Analysis & Ratio Modeling]
- Apply standard accounting & corporate finance principles.
- Automatically format currency ($#,##0) and percentage (0.0%) values using format_range.
- Double-check sign conventions (Revenue positive, Expenses negative or positive depending on context).
- Explain key ratio formulas clearly in plain language.`,
  },
  {
    id: 'formula_expert',
    name: 'Advanced Excel Formula Architect',
    category: 'Formulas',
    description: 'Write robust, error-safe Excel formulas (XLOOKUP, INDEX-MATCH, SUMIFS, LET, LAMBDA, IFERROR).',
    isBuiltIn: true,
    instructions: `[SPECIALIZED SKILL: Advanced Excel Formula Architect]
- Prefer modern dynamic array formulas (XLOOKUP, FILTER, SORT, UNIQUE) when available.
- Always wrap lookup formulas in IFERROR or IFNA to prevent ugly #N/A or #REF! errors.
- Ensure cell references maintain relative/absolute ($A$1) logic appropriately when filled.`,
  },
];

const CUSTOM_SKILLS_STORAGE_KEY = 'llmgateway_excel_custom_skills';

export function getCustomSkills(): Skill[] {
  try {
    const raw = localStorage.getItem(CUSTOM_SKILLS_STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (err) {
    console.error('Failed to load custom skills:', err);
    return [];
  }
}

export function saveCustomSkills(skills: Skill[]): void {
  try {
    localStorage.setItem(CUSTOM_SKILLS_STORAGE_KEY, JSON.stringify(skills));
  } catch (err) {
    console.error('Failed to save custom skills:', err);
  }
}

export function getAllSkills(): Skill[] {
  return [...BUILTIN_SKILLS, ...getCustomSkills()];
}

export function getActiveSkillsInstructions(activeSkillIds: string[]): string {
  const all = getAllSkills();
  const active = all.filter((s) => activeSkillIds.includes(s.id));
  if (active.length === 0) return '';
  return `\n\n=== ACTIVE AGENT SKILLS ===\n` + active.map((s) => s.instructions).join('\n\n');
}
