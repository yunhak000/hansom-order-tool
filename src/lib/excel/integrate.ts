import ExcelJS from "exceljs";
import { TStandardRow } from "@/lib/types";
import { readWorkbook, getFirstSheet, readHeadersFromRow1 } from "./read";

// 템플릿의 헤더명으로 열 인덱스 찾기
const buildColIndex = (headers: string[]) => {
  const map = new Map<string, number>();
  headers.forEach((h, idx) => map.set(h, idx + 1));
  return map;
};

export const buildIntegrationWorkbook = async (
  templateArrayBuffer: ArrayBuffer,
  standardRows: TStandardRow[],
): Promise<ArrayBuffer> => {
  const wb = await readWorkbook(templateArrayBuffer);
  const ws = getFirstSheet(wb);

  const headers = readHeadersFromRow1(ws);
  const col = buildColIndex(headers);

  // 기존 데이터 영역(2행~)을 지우고 새로 쓰기 (서식은 템플릿 2행을 복제하는 방식이 안정적)
  // 템플릿이 “2행에 예시 데이터”가 있다면, 그 행 스타일을 복제해서 사용
  const baseRow = ws.getRow(2);

  // 기존 2행~ 끝까지 제거
  if (ws.rowCount >= 2) {
    ws.spliceRows(2, ws.rowCount - 1);
  }

  const writeCell = (row: ExcelJS.Row, header: string, value: any) => {
    const c = col.get(header);
    if (!c) return;
    row.getCell(c).value = value ?? "";
  };

  standardRows.forEach((s, i) => {
    const rIndex = 2 + i;
    ws.insertRow(rIndex, []); // 빈 행 삽입
    const row = ws.getRow(rIndex);

    // baseRow 스타일 복제
    row.height = baseRow.height;
    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      const baseCell = baseRow.getCell(colNum);
      cell.style = { ...baseCell.style };
      cell.numFmt = baseCell.numFmt;
      cell.border = baseCell.border;
      cell.alignment = baseCell.alignment;
      cell.font = baseCell.font;
      cell.fill = baseCell.fill;
    });

    writeCell(row, "주문일시", s.orderedAt ?? "");
    writeCell(row, "상품명", s.productName);
    writeCell(row, "수량", s.quantity);
    writeCell(row, "수취인명", s.receiverName);
    writeCell(row, "수취인연락처1", s.receiverPhone);
    writeCell(row, "우편번호", s.zipCode ?? "");
    writeCell(row, "통합배송지", s.address);
    writeCell(row, "배송메세지", s.message ?? "");
    writeCell(row, "상품주문번호", s.orderKey);
    writeCell(row, "운송장번호", ""); // 통합 단계에서는 항상 비움
    writeCell(row, "구매자명", s.buyerName);
    writeCell(row, "구매자연락처", s.buyerPhone);
    writeCell(row, "어드민용 구매자명", s.adminBuyerName);
    writeCell(row, "어드민용 구매자연락처", s.adminBuyerPhone);

    row.commit();
  });

  const out = await wb.xlsx.writeBuffer();
  return out as ArrayBuffer;
};
