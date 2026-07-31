import { OpenAIToolDefinition, ToolExecutionResult } from '../types';

export const EXCEL_TOOLS_SCHEMA: OpenAIToolDefinition[] = [
  {
    type: 'function',
    function: {
      name: 'get_workbook_overview',
      description: 'Get an overview of the active workbook, active worksheet name, list of worksheets, tables, and currently selected cell address.',
      parameters: {
        type: 'object',
        properties: {},
        required: [],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'read_range',
      description: 'Read cell values and formulas from a specific cell range address (e.g. "A1:C10" or "Sheet1!A1:B5").',
      parameters: {
        type: 'object',
        properties: {
          address: {
            type: 'string',
            description: 'The cell range address, e.g. "A1:D10" or "Sheet2!B2:E20". If empty, reads current selection.',
          },
        },
        required: [],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'write_cells',
      description: 'Write values or formulas to a range of cells starting at target address.',
      parameters: {
        type: 'object',
        properties: {
          address: {
            type: 'string',
            description: 'Target cell range address, e.g. "A1:B2" or "Sheet1!C5".',
          },
          values: {
            type: 'array',
            items: {
              type: 'array',
              items: {},
            },
            description: '2D array of values or formulas (e.g. [["Name", "Score"], ["Alice", 95], ["Bob", "=SUM(B2:B3)"]]).',
          },
        },
        required: ['address', 'values'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'format_range',
      description: 'Format a cell range (fill color, font color, bold, number format, auto-fit columns).',
      parameters: {
        type: 'object',
        properties: {
          address: {
            type: 'string',
            description: 'Cell range address to format, e.g. "A1:D1".',
          },
          fillColor: {
            type: 'string',
            description: 'Hex background fill color (e.g. "#1E3A8A" or "yellow").',
          },
          fontColor: {
            type: 'string',
            description: 'Hex font color (e.g. "#FFFFFF" or "black").',
          },
          bold: {
            type: 'boolean',
            description: 'Set font bold state.',
          },
          numberFormat: {
            type: 'string',
            description: 'Excel number format string (e.g. "$#,##0.00", "0.0%", "YYYY-MM-DD").',
          },
          autofit: {
            type: 'boolean',
            description: 'Auto-fit column widths for the range.',
          },
        },
        required: ['address'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'clear_range',
      description: 'Clear values, formulas, or formats from a range of cells.',
      parameters: {
        type: 'object',
        properties: {
          address: {
            type: 'string',
            description: 'Cell range address to clear.',
          },
          clearType: {
            type: 'string',
            enum: ['contents', 'formats', 'all'],
            description: 'What to clear: "contents" (values/formulas), "formats" (styling), or "all". Default "all".',
          },
        },
        required: ['address'],
      },
    },
  },
];

export async function executeExcelTool(
  name: string,
  args: Record<string, unknown>
): Promise<ToolExecutionResult> {
  if (typeof Office === 'undefined' || !Office.context) {
    return {
      success: false,
      message: 'Office.js is not loaded or not running inside Excel environment.',
    };
  }

  try {
    switch (name) {
      case 'get_workbook_overview': {
        return await Excel.run(async (context) => {
          const activeSheet = context.workbook.worksheets.getActiveWorksheet();
          const worksheets = context.workbook.worksheets;
          const selection = context.workbook.getSelectedRange();

          activeSheet.load('name');
          worksheets.load('items/name');
          selection.load('address');

          await context.sync();

          const sheetNames = worksheets.items.map((s) => s.name);

          return {
            success: true,
            message: `Active sheet: "${activeSheet.name}", Worksheets: [${sheetNames.join(', ')}], Selection: ${selection.address}`,
            data: {
              activeSheet: activeSheet.name,
              sheets: sheetNames,
              selectionAddress: selection.address,
            },
          };
        });
      }

      case 'read_range': {
        const addressArg = (args.address as string) || '';
        return await Excel.run(async (context) => {
          let range: Excel.Range;
          if (addressArg.trim()) {
            range = context.workbook.worksheets.getActiveWorksheet().getRange(addressArg);
          } else {
            range = context.workbook.getSelectedRange();
          }

          range.load(['address', 'values', 'formulas', 'text', 'rowCount', 'columnCount']);
          await context.sync();

          return {
            success: true,
            message: `Successfully read range ${range.address} (${range.rowCount}x${range.columnCount})`,
            data: {
              address: range.address,
              values: range.values,
              formulas: range.formulas,
              text: range.text,
            },
          };
        });
      }

      case 'write_cells': {
        const address = args.address as string;
        const values = args.values as unknown[][];
        if (!address || !Array.isArray(values)) {
          return { success: false, message: 'Invalid arguments: "address" and "values" array are required.' };
        }

        return await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const range = sheet.getRange(address);
          range.values = values;
          range.load(['address', 'rowCount', 'columnCount']);
          await context.sync();

          return {
            success: true,
            message: `Successfully wrote ${values.length} rows to ${range.address}`,
            data: { address: range.address },
          };
        });
      }

      case 'format_range': {
        const address = args.address as string;
        const fillColor = args.fillColor as string | undefined;
        const fontColor = args.fontColor as string | undefined;
        const bold = args.bold as boolean | undefined;
        const numberFormat = args.numberFormat as string | undefined;
        const autofit = args.autofit as boolean | undefined;

        if (!address) {
          return { success: false, message: 'Invalid argument: "address" is required.' };
        }

        return await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const range = sheet.getRange(address);

          if (fillColor) range.format.fill.color = fillColor;
          if (fontColor) range.format.font.color = fontColor;
          if (typeof bold === 'boolean') range.format.font.bold = bold;
          if (numberFormat) range.numberFormat = [[numberFormat]];
          if (autofit) range.format.autofitColumns();

          range.load('address');
          await context.sync();

          return {
            success: true,
            message: `Successfully formatted range ${range.address}`,
            data: { address: range.address },
          };
        });
      }

      case 'clear_range': {
        const address = args.address as string;
        const clearType = (args.clearType as string) || 'all';

        if (!address) {
          return { success: false, message: 'Invalid argument: "address" is required.' };
        }

        return await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const range = sheet.getRange(address);

          if (clearType === 'contents') {
            range.clear(Excel.ClearApplyTo.contents);
          } else if (clearType === 'formats') {
            range.clear(Excel.ClearApplyTo.formats);
          } else {
            range.clear(Excel.ClearApplyTo.all);
          }

          range.load('address');
          await context.sync();

          return {
            success: true,
            message: `Cleared (${clearType}) for range ${range.address}`,
            data: { address: range.address },
          };
        });
      }

      default:
        return {
          success: false,
          message: `Unknown tool name: ${name}`,
        };
    }
  } catch (err: unknown) {
    const errorMsg = err instanceof Error ? err.message : String(err);
    console.error(`Error executing tool ${name}:`, err);
    return {
      success: false,
      message: `Tool execution failed: ${errorMsg}`,
    };
  }
}
