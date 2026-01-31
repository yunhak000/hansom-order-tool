import { TChannel, THansomResultMap } from "@/lib/types";
import { readWorkbook, getFirstSheet, readHeadersAuto } from "./read";

const channelOrderKeyHeader: Record<TChannel, string> = {
  NAVER: "상품주문번호",
  TOSS: "주문상품번호",
  COUPANG: "주문번호",
  MANDARINSPOON: "고객주문번호",
};

const channelTrackingHeader: Record<TChannel, string> = {
  NAVER: "송장번호",
  TOSS: "송장번호",
  COUPANG: "운송장번호",
  MANDARINSPOON: "운송장번호",
};

const findHeaderCol = (headers: string[], target: string) => {
  const exact = headers.indexOf(target);
  if (exact !== -1) return exact + 1;

  const normalizedTarget = target.replace(/\s/g, "");
  const partial = headers.findIndex((h) =>
    String(h ?? "")
      .replace(/\s/g, "")
      .includes(normalizedTarget),
  );

  return partial !== -1 ? partial + 1 : -1;
};

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const setTextCell = (cell: any, value: unknown) => {
  cell.value = value == null ? "" : String(value);
  cell.numFmt = "@";
};

// ✅ 날짜 컬럼 표시 형식 강제
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const setDateColumnFormat = (ws: any, colIndex: number) => {
  ws.getColumn(colIndex).numFmt = "yyyy-mm-dd hh:mm";
};

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

export const fillTrackingToOriginal = async (
  channel: TChannel,
  originalArrayBuffer: ArrayBuffer,
  hansomMap: THansomResultMap,
): Promise<ArrayBuffer> => {
  const wb = await readWorkbook(originalArrayBuffer);
  const ws = getFirstSheet(wb);

  if (channel === "NAVER") {
    ws.name = "발송처리";
  }

  let { headers, headerRowIndex } = readHeadersAuto(ws);

  // NAVER: 디스크립션 제거 후 재탐색
  if (channel === "NAVER" && headerRowIndex > 1) {
    ws.spliceRows(1, headerRowIndex - 1);
    ({ headers, headerRowIndex } = readHeadersAuto(ws));
  }

  const keyHeader = channelOrderKeyHeader[channel];
  const trackingHeader = channelTrackingHeader[channel];

  const keyCol = findHeaderCol(headers, keyHeader);
  const trackCol = findHeaderCol(headers, trackingHeader);

  if (keyCol <= 0)
    throw new Error(
      `${channel} 엑셀에서 주문번호 컬럼(${keyHeader})을 찾지 못했어요.`,
    );
  if (trackCol <= 0)
    throw new Error(
      `${channel} 엑셀에서 운송장 컬럼(${trackingHeader})을 찾지 못했어요.`,
    );

  ws.getColumn(keyCol).numFmt = "@";
  ws.getColumn(trackCol).numFmt = "@";

  // ✅ NAVER 날짜 컬럼들이 숫자로 보이는 문제 해결
  if (channel === "NAVER") {
    const dateHeaders = ["발주확인일", "발송기한"];
    for (const h of dateHeaders) {
      const c = findHeaderCol(headers, h);
      if (c > 0) setDateColumnFormat(ws, c);
    }
  }

  for (let r = headerRowIndex + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);

    const key = String(row.getCell(keyCol).value ?? "").trim();
    if (!key) continue;

    const tracking = hansomMap.get(key);
    if (tracking) {
      setTextCell(row.getCell(trackCol), tracking);
    }

    row.commit();
  }

  const out = await wb.xlsx.writeBuffer();
  return toArrayBuffer(out);
};
