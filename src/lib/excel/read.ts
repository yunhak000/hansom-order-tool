import ExcelJS from "exceljs";

const normHeader = (v: unknown) =>
  String(v ?? "")
    .trim()
    .replace(/\s+/g, " ");

const HEADER_KEYWORDS = [
  // 주문번호 계열
  "상품주문번호",
  "주문상품번호",
  "주문번호",
  "고객주문번호",
  "거래처 주문번호",

  // 운송장/송장
  "운송장번호",
  "송장번호",

  // 네이버/통합 쪽에서 자주 등장하는 필드
  "주문일시",
  "상품명",
  "수량",
  "수취인명",
  "구매자명",
  "통합배송지",
];

export const readHeadersFromRow = (ws: ExcelJS.Worksheet, rowIndex: number) => {
  const row = ws.getRow(rowIndex);
  const headers: string[] = [];

  // ExcelJS는 1-based index
  for (let c = 1; c <= row.cellCount; c++) {
    const v = row.getCell(c).value;
    const h = normHeader(v);
    headers.push(h);
  }

  // 끝에 빈값이 길게 붙는 경우 잘라냄
  while (headers.length && !headers[headers.length - 1]) headers.pop();

  return headers;
};

export const findHeaderRowIndex = (ws: ExcelJS.Worksheet, scanRows = 20) => {
  let bestRow = 1;
  let bestScore = -1;

  for (let r = 1; r <= Math.min(scanRows, ws.rowCount); r++) {
    const headers = readHeadersFromRow(ws, r);
    if (!headers.length) continue;

    const score = headers.reduce((acc, h) => {
      if (!h) return acc;
      return acc + (HEADER_KEYWORDS.includes(h) ? 1 : 0);
    }, 0);

    // 최소 2개 이상 키워드가 매칭되면 헤더 후보로 인정
    if (score > bestScore && score >= 2) {
      bestScore = score;
      bestRow = r;
    }
  }

  return bestRow;
};

export const readHeadersAuto = (ws: ExcelJS.Worksheet, scanRows = 20) => {
  const headerRowIndex = findHeaderRowIndex(ws, scanRows);
  const headers = readHeadersFromRow(ws, headerRowIndex);
  return { headers, headerRowIndex };
};

export const readWorkbook = async (arrayBuffer: ArrayBuffer) => {
  const wb = new ExcelJS.Workbook();

  try {
    await wb.xlsx.load(arrayBuffer);
  } catch (e) {
    // 암호/손상/특수 포맷 등으로 파싱 실패 시
    throw new Error(
      "엑셀을 읽을 수 없어요. (네이버 파일이라면 비밀번호를 해제한 뒤 다시 저장해서 업로드해주세요)",
    );
  }

  // 워크시트가 0개로 인식되는 경우도 암호/특수 포맷에서 자주 발생
  if (!wb.worksheets || wb.worksheets.length === 0) {
    throw new Error(
      "엑셀 시트를 찾지 못했어요. (네이버 파일이라면 비밀번호를 해제한 뒤 다시 저장해서 업로드해주세요)",
    );
  }

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

export const readRowsAsObjects = (
  ws: ExcelJS.Worksheet,
  headers: string[],
  headerRowIndex = 1,
) => {
  const out: Record<string, unknown>[] = [];

  for (let r = headerRowIndex + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const obj: Record<string, unknown> = {};

    headers.forEach((h, i) => {
      if (!h) return;
      obj[h] = row.getCell(i + 1).value;
    });

    // 완전 빈 행은 스킵
    const hasAny = Object.values(obj).some(
      (v) => v != null && String(v).trim(),
    );
    if (hasAny) out.push(obj);
  }

  return out;
};
