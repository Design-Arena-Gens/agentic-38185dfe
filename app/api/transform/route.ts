import { NextRequest, NextResponse } from 'next/server';
import ExcelJS from 'exceljs';
import { parseInstruction, type Action } from '@/src/agent/parseInstruction';

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

export async function POST(req: NextRequest) {
  try {
    const form = await req.formData();
    const file = form.get('file');
    const instruction = String(form.get('instruction') || '').trim();

    if (!(file instanceof File)) {
      return new NextResponse('No file uploaded', { status: 400 });
    }
    if (!instruction) {
      return new NextResponse('Instruction required', { status: 400 });
    }

    const arrayBuffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(Buffer.from(arrayBuffer));

    const actions = parseInstruction(instruction, workbook);

    for (const action of actions) {
      await applyAction(workbook, action);
    }

    const outBuffer = await workbook.xlsx.writeBuffer();

    return new NextResponse(outBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="updated.xlsx"'
      }
    });
  } catch (err: any) {
    console.error(err);
    return new NextResponse(err?.message || 'Server error', { status: 500 });
  }
}

async function applyAction(workbook: ExcelJS.Workbook, action: Action): Promise<void> {
  switch (action.type) {
    case 'rename_sheet': {
      const sheet = workbook.getWorksheet(action.sheetOld) || workbook.worksheets[0];
      if (sheet) sheet.name = action.sheetNew;
      break;
    }
    case 'rename_column': {
      const sheet = getSheet(workbook, action.sheetName);
      if (!sheet) break;
      const headerRow = sheet.getRow(1);
      for (let col = 1; col <= headerRow.cellCount; col++) {
        const v = String(headerRow.getCell(col).value || '').trim();
        if (normalize(v) === normalize(action.columnOld)) {
          headerRow.getCell(col).value = action.columnNew;
          headerRow.commit?.();
          break;
        }
      }
      break;
    }
    case 'add_column_sum': {
      const sheet = getSheet(workbook, action.sheetName);
      if (!sheet) break;
      const headerRow = sheet.getRow(1);
      const colA = findColumnIndexByHeader(sheet, action.colA);
      const colB = findColumnIndexByHeader(sheet, action.colB);
      if (!colA || !colB) break;
      const newColIndex = headerRow.cellCount + 1;
      headerRow.getCell(newColIndex).value = action.newColumn;
      headerRow.commit?.();
      for (let r = 2; r <= sheet.rowCount; r++) {
        const cellA = Number(sheet.getRow(r).getCell(colA).value ?? 0);
        const cellB = Number(sheet.getRow(r).getCell(colB).value ?? 0);
        sheet.getRow(r).getCell(newColIndex).value = cellA + cellB;
        sheet.getRow(r).commit?.();
      }
      break;
    }
    case 'delete_column': {
      const sheet = getSheet(workbook, action.sheetName);
      if (!sheet) break;
      const idx = findColumnIndexByHeader(sheet, action.columnName);
      if (idx) sheet.spliceColumns(idx, 1);
      break;
    }
    case 'filter_rows': {
      const sheet = getSheet(workbook, action.sheetName);
      if (!sheet) break;
      const idx = findColumnIndexByHeader(sheet, action.columnName);
      if (!idx) break;
      const op = action.operator; // '>', '>=', '<', '<=', '==', '!='
      const target = action.value;
      const keepRows: number[] = [1];
      for (let r = 2; r <= sheet.rowCount; r++) {
        const valRaw = sheet.getRow(r).getCell(idx).value as any;
        const val = typeof valRaw === 'number' ? valRaw : Number(String(valRaw).replace(/,/g, ''));
        const cond = compare(val, target, op);
        if (cond) keepRows.push(r);
      }
      // Remove rows not in keepRows (from bottom to top)
      for (let r = sheet.rowCount; r >= 2; r--) {
        if (!keepRows.includes(r)) sheet.spliceRows(r, 1);
      }
      break;
    }
    case 'sort_by': {
      const sheet = getSheet(workbook, action.sheetName);
      if (!sheet) break;
      const idx = findColumnIndexByHeader(sheet, action.columnName);
      if (!idx) break;
      const rows: Array<any[]> = [];
      for (let r = 2; r <= sheet.rowCount; r++) {
        rows.push(sheet.getRow(r).values as any[]);
      }
      rows.sort((a, b) => {
        const av = a[idx];
        const bv = b[idx];
        const an = typeof av === 'number' ? av : Number(av);
        const bn = typeof bv === 'number' ? bv : Number(bv);
        if (!isNaN(an) && !isNaN(bn)) return action.direction === 'asc' ? an - bn : bn - an;
        const as = String(av ?? '');
        const bs = String(bv ?? '');
        return action.direction === 'asc' ? as.localeCompare(bs) : bs.localeCompare(as);
      });
      // write back
      for (let r = 2; r <= sheet.rowCount; r++) sheet.spliceRows(r, 1);
      rows.forEach((vals, i) => {
        const row = sheet.getRow(2 + i);
        for (let c = 1; c < vals.length; c++) row.getCell(c).value = vals[c];
        row.commit?.();
      });
      break;
    }
    case 'set_value_where': {
      const sheet = getSheet(workbook, action.sheetName);
      if (!sheet) break;
      const idxTarget = findColumnIndexByHeader(sheet, action.targetColumn);
      const idxCond = findColumnIndexByHeader(sheet, action.conditionColumn);
      if (!idxTarget || !idxCond) break;
      for (let r = 2; r <= sheet.rowCount; r++) {
        const raw = sheet.getRow(r).getCell(idxCond).value as any;
        const pass = compare(String(raw), String(action.conditionValue), action.operator);
        if (pass) {
          sheet.getRow(r).getCell(idxTarget).value = action.value;
          sheet.getRow(r).commit?.();
        }
      }
      break;
    }
    default:
      break;
  }
}

function getSheet(workbook: ExcelJS.Workbook, name?: string) {
  if (name) return workbook.getWorksheet(name) || workbook.worksheets[0];
  return workbook.worksheets[0];
}

function findColumnIndexByHeader(sheet: ExcelJS.Worksheet, header: string): number | undefined {
  const headerRow = sheet.getRow(1);
  for (let c = 1; c <= headerRow.cellCount; c++) {
    const v = String(headerRow.getCell(c).value || '').trim();
    if (normalize(v) === normalize(header)) return c;
  }
  return undefined;
}

function normalize(s: string) {
  return s.toLowerCase().replace(/\s+/g, ' ').trim();
}

function compare(a: any, b: any, op: string): boolean {
  // Try numeric compare, else string compare
  const an = Number(a);
  const bn = Number(b);
  const hasNum = !isNaN(an) && !isNaN(bn);
  if (hasNum) {
    switch (op) {
      case '>': return an > bn;
      case '>=': return an >= bn;
      case '<': return an < bn;
      case '<=': return an <= bn;
      case '!=': return an !== bn;
      case '==': default: return an === bn;
    }
  } else {
    const as = String(a ?? '');
    const bs = String(b ?? '');
    switch (op) {
      case '!=': return as !== bs;
      case '==': default: return as === bs;
    }
  }
}
