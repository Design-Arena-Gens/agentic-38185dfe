import type ExcelJS from 'exceljs';

export type Action =
  | { type: 'rename_sheet'; sheetOld: string; sheetNew: string }
  | { type: 'rename_column'; sheetName?: string; columnOld: string; columnNew: string }
  | { type: 'add_column_sum'; sheetName?: string; colA: string; colB: string; newColumn: string }
  | { type: 'delete_column'; sheetName?: string; columnName: string }
  | { type: 'filter_rows'; sheetName?: string; columnName: string; operator: '>'|'<'|'>='|'<='|'=='|'!='; value: number }
  | { type: 'sort_by'; sheetName?: string; columnName: string; direction: 'asc'|'desc' }
  | { type: 'set_value_where'; sheetName?: string; targetColumn: string; operator: '=='|'!='; conditionColumn: string; conditionValue: string; value: string };

function norm(s: string) { return s.toLowerCase().trim(); }

function extractSheetMention(text: string, workbook: ExcelJS.Workbook): string | undefined {
  // Try to find an existing sheet name mentioned in text
  const lower = text.toLowerCase();
  for (const ws of workbook.worksheets) {
    const name = ws.name;
    if (!name) continue;
    if (lower.includes(name.toLowerCase())) return name;
  }
  return undefined;
}

export function parseInstruction(instruction: string, workbook: ExcelJS.Workbook): Action[] {
  const text = instruction.replace(/\s+/g, ' ').trim();
  const lower = text.toLowerCase();
  const actions: Action[] = [];

  // 1) Rename sheet: "sales sheet ka naam revenue rakho" or "rename sheet Sales to Revenue"
  {
    const m = lower.match(/(?:sheet|sheeet|sheets?)\s+(?:ka\s+naam|name)?\s*(\w[\w\s-]*)\s+(?:ko|to|rakho|banao)\s+(\w[\w\s-]*)/);
    if (m) {
      const from = capitalize(m[1]);
      const to = capitalize(m[2]);
      actions.push({ type: 'rename_sheet', sheetOld: from, sheetNew: to });
      return actions;
    }
  }

  // 2) Rename column: "Price column ka naam Cost rakho" or "rename column Price to Cost"
  {
    const m = lower.match(/(?:column|col|stambh|??????)\s+(\w[\w\s-]*)\s+(?:ka\s+naam|name)?\s*(?:ko|to|rakho|banao)\s+(\w[\w\s-]*)/);
    if (m) {
      const colOld = capitalize(m[1]);
      const colNew = capitalize(m[2]);
      actions.push({ type: 'rename_column', sheetName: extractSheetMention(lower, workbook), columnOld: colOld, columnNew: colNew });
      return actions;
    }
  }

  // 3) Add column sum: "Total column add karo jo Price + Tax ho"
  {
    const m = lower.match(/(\w[\w\s-]*)\s+column\s+(?:add|create|banao)\s+(?:karo|kar do)?\s*(?:jo|which)?\s*(\w[\w\s-]*)\s*\+\s*(\w[\w\s-]*)/);
    if (m) {
      const newCol = capitalize(m[1]);
      const colA = capitalize(m[2]);
      const colB = capitalize(m[3]);
      actions.push({ type: 'add_column_sum', sheetName: extractSheetMention(lower, workbook), colA, colB, newColumn: newCol });
      return actions;
    }
  }

  // 4) Delete column: "delete column Discount" or "Discount column hatao"
  {
    const m = lower.match(/(?:delete|remove|hatao|hatado)\s+(?:column|col)\s+(\w[\w\s-]*)|^(\w[\w\s-]*)\s+(?:column|col)\s+(?:delete|remove|hatao)/);
    if (m) {
      const name = capitalize(m[1] || m[2]);
      if (name) {
        actions.push({ type: 'delete_column', sheetName: extractSheetMention(lower, workbook), columnName: name });
        return actions;
      }
    }
  }

  // 5) Filter rows keep only where condition e.g. "sirf woh rows jahan Quantity > 10 rakho" or "keep rows where Quantity > 10"
  {
    const m = lower.match(/(?:where|jahan)\s+(\w[\w\s-]*)\s*(>=|<=|==|!=|>|<)\s*([0-9]+(?:\.[0-9]+)?)/);
    if (m && (lower.includes('sirf') || lower.includes('keep') || lower.includes('rakh'))) {
      const col = capitalize(m[1]);
      const op = m[2] as Action['operator'];
      const value = Number(m[3]);
      actions.push({ type: 'filter_rows', sheetName: extractSheetMention(lower, workbook), columnName: col, operator: op, value });
      return actions;
    }
  }

  // 6) Sort by column: "sort by Amount desc" or "Amount ke hisaab se sort desc"
  {
    const m = lower.match(/sort\s+(?:by\s+)?(\w[\w\s-]*)(?:\s+(asc|desc))?/);
    if (m) {
      const col = capitalize(m[1]);
      const dir = (m[2] as 'asc'|'desc') || 'asc';
      actions.push({ type: 'sort_by', sheetName: extractSheetMention(lower, workbook), columnName: col, direction: dir });
      return actions;
    }
  }

  // 7) Set value where: "Status ko 'Done' set karo jahan Type == 'A'"
  {
    const m = lower.match(/(\w[\w\s-]*)\s+(?:ko|to)?\s*'([^']+)'\s*(?:set|rakho)\s*(?:karo|do)?\s*(?:jahan|where)\s+(\w[\w\s-]*)\s*(==|!=)\s*'([^']+)'/);
    if (m) {
      const target = capitalize(m[1]);
      const value = m[2];
      const condCol = capitalize(m[3]);
      const op = m[4] as '=='|'!=';
      const condVal = m[5];
      actions.push({ type: 'set_value_where', sheetName: extractSheetMention(lower, workbook), targetColumn: target, operator: op, conditionColumn: condCol, conditionValue: condVal, value });
      return actions;
    }
  }

  // Default: no-op; return empty so server returns original workbook
  return actions;
}

function capitalize(s: string) {
  return s
    .split(' ')
    .map((w) => (w.length ? w[0].toUpperCase() + w.slice(1) : w))
    .join(' ')
    .trim();
}
