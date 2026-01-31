import { TChannel, THansomResultMap } from "@/lib/types";
import { readWorkbook, getFirstSheet, readHeadersFromRow1 } from "./read";

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

export const fillTrackingToOriginal = async (
  channel: TChannel,
  originalArrayBuffer: ArrayBuffer,
  hansomMap: THansomResultMap,
): Promise<ArrayBuffer> => {
  const wb = await readWorkbook(originalArrayBuffer);
  const ws = getFirstSheet(wb);

  const headers = readHeadersFromRow1(ws);
  const keyHeader = channelOrderKeyHeader[channel];
  const trackingHeader = channelTrackingHeader[channel];

  const keyCol = headers.indexOf(keyHeader) + 1;
  const trackCol = headers.indexOf(trackingHeader) + 1;

  if (keyCol <= 0)
    throw new Error(
      `${channel} 엑셀에서 주문번호 컬럼(${keyHeader})을 찾지 못했어요.`,
    );
  if (trackCol <= 0)
    throw new Error(
      `${channel} 엑셀에서 운송장 컬럼(${trackingHeader})을 찾지 못했어요.`,
    );

  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const key = String(row.getCell(keyCol).value ?? "").trim();
    if (!key) continue;

    const tracking = hansomMap.get(key);
    if (tracking) {
      // 텍스트로 강제
      row.getCell(trackCol).value = String(tracking);
    }
    row.commit();
  }

  const out = await wb.xlsx.writeBuffer();
  return out as ArrayBuffer;
};
