import ExcelJS from "exceljs";
import { TStandardRow } from "@/lib/types";
import { readWorkbook, getFirstSheet, readHeadersFromRow1 } from "./read";

const buildColIndex = (headers: string[]) => {
  const map = new Map<string, number>();
  headers.forEach((h, idx) => map.set(h, idx + 1));
  return map;
};

const dedupeByOrderKey = <T extends { channel: string; orderKey: string }>(
  rows: T[],
) => {
  const map = new Map<string, T>();
  for (const r of rows) {
    const key = `${r.channel}:${r.orderKey}`;
    if (!map.has(key)) map.set(key, r);
  }
  return Array.from(map.values());
};

// ✅ ArrayBuffer로 강제 변환 (SharedArrayBuffer 대응)
const toArrayBuffer = (out: unknown): ArrayBuffer => {
  if (out instanceof ArrayBuffer) return out;

  if (
    typeof SharedArrayBuffer !== "undefined" &&
    out instanceof SharedArrayBuffer
  ) {
    const ab = new ArrayBuffer(out.byteLength);
    new Uint8Array(ab).set(new Uint8Array(out));
    return ab;
  }

  if (out instanceof Uint8Array) {
    const ab = new ArrayBuffer(out.byteLength);
    new Uint8Array(ab).set(out);
    return ab;
  }

  throw new Error("엑셀 버퍼 변환에 실패했어요.");
};

export const buildIntegrationWorkbook = async (
  templateArrayBuffer: ArrayBuffer,
  standardRows: TStandardRow[],
): Promise<ArrayBuffer> => {
  const wb = await readWorkbook(templateArrayBuffer);
  const ws = getFirstSheet(wb);

  const headers = readHeadersFromRow1(ws);
  const col = buildColIndex(headers);

  const rows = dedupeByOrderKey(standardRows);

  const baseRow = ws.getRow(2);

  if (ws.rowCount >= 2) {
    ws.spliceRows(2, ws.rowCount - 1);
  }

  const toCellValue = (v: unknown): ExcelJS.CellValue => {
    if (v == null) return "";

    // ExcelJS가 허용하는 기본 타입들
    if (typeof v === "string") return v;
    if (typeof v === "number") return v;
    if (typeof v === "boolean") return v;
    if (v instanceof Date) return v;

    // 나머지는 안전하게 문자열로 변환
    return String(v);
  };

  const writeCell = (row: ExcelJS.Row, header: string, value: unknown) => {
    const c = col.get(header);
    if (!c) return;

    row.getCell(c).value = toCellValue(value);
  };

  rows.forEach((s, i) => {
    const rIndex = 2 + i;
    ws.insertRow(rIndex, []);
    const row = ws.getRow(rIndex);

    row.height = baseRow.height;

    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      const baseCell = baseRow.getCell(colNum);
      cell.numFmt = baseCell.numFmt;
      cell.border = baseCell.border;
      cell.alignment = baseCell.alignment;
      cell.font = baseCell.font;
      cell.fill = baseCell.fill;
    });

    writeCell(row, "주문일시", String(s.orderedAt ?? ""));
    writeCell(row, "상품명", s.productName);
    writeCell(row, "수량", s.quantity);
    writeCell(row, "수취인명", s.receiverName);
    writeCell(row, "수취인연락처1", s.receiverPhone);
    writeCell(row, "우편번호", s.zipCode ?? "");
    writeCell(row, "통합배송지", s.address);
    writeCell(row, "배송메세지", s.message ?? "");
    writeCell(row, "거래처주문번호", s.orderKey);
    writeCell(row, "운송장번호", "");
    writeCell(row, "구매자명", s.buyerName);
    writeCell(row, "구매자연락처", s.buyerPhone);
    writeCell(row, "어드민용 구매자명", s.adminBuyerName);
    writeCell(row, "어드민용 구매자연락처", s.adminBuyerPhone);

    row.commit();
  });

  const out = await wb.xlsx.writeBuffer();
  return toArrayBuffer(out);
};
