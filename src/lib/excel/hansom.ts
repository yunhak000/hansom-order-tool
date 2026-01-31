import {
  readWorkbook,
  getFirstSheet,
  readHeadersFromRow1,
  readRowsAsObjects,
} from "./read";
import { THansomResultMap } from "@/lib/types";

export const buildHansomMap = async (
  arrayBuffer: ArrayBuffer,
): Promise<THansomResultMap> => {
  const wb = await readWorkbook(arrayBuffer);
  const ws = getFirstSheet(wb);
  const headers = readHeadersFromRow1(ws);
  const rows = readRowsAsObjects(ws, headers);

  // 키 컬럼: 거래처 주문번호 (없으면 상품주문번호로 폴백)
  const keyField = headers.includes("거래처 주문번호")
    ? "거래처 주문번호"
    : headers.includes("상품주문번호")
      ? "상품주문번호"
      : null;

  if (!keyField)
    throw new Error(
      "한섬누리 결과 엑셀에서 주문번호 컬럼을 찾지 못했어요. (거래처 주문번호/상품주문번호)",
    );

  if (!headers.includes("운송장번호"))
    throw new Error("한섬누리 결과 엑셀에서 운송장번호 컬럼을 찾지 못했어요.");

  const map: THansomResultMap = new Map();
  rows.forEach((r) => {
    const k = String(r[keyField] ?? "").trim();
    const t = String(r["운송장번호"] ?? "").trim();
    if (k) map.set(k, t);
  });

  return map;
};
