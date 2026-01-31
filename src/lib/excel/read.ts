import ExcelJS from "exceljs";

export const readWorkbook = async (arrayBuffer: ArrayBuffer) => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);
  return wb;
};

export const getFirstSheet = (wb: ExcelJS.Workbook) => {
  const ws = wb.worksheets[0];
  if (!ws) throw new Error("엑셀 시트를 찾지 못했어요.");
  return ws;
};

export const readHeadersFromRow1 = (ws: ExcelJS.Worksheet) => {
  const row1 = ws.getRow(1);
  const headers: string[] = [];
  row1.eachCell({ includeEmpty: true }, (cell, col) => {
    headers[col - 1] = String(cell.value ?? "").trim();
  });
  // trailing empty 제거
  return headers.filter((h) => h);
};

export const readRowsAsObjects = (ws: ExcelJS.Worksheet, headers: string[]) => {
  const rows: Record<string, unknown>[] = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const obj: Record<string, unknown> = {};
    let empty = true;

    headers.forEach((h, i) => {
      const cell = row.getCell(i + 1);
      const v = cell.value;

      if (v !== null && v !== undefined && String(v).trim() !== "")
        empty = false;
      obj[h] =
        typeof v === "object" && v && "text" in (v as any)
          ? (v as any).text
          : v;
    });

    if (!empty) rows.push(obj);
  }
  return rows;
};
