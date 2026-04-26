import { TChannel, TStandardRow } from "@/lib/types";

const normName = (v: unknown) =>
  String(v ?? "")
    .trim()
    .replace(/\s+/g, " ");

const normPhone = (v: unknown) => String(v ?? "").trim();

const toNumber = (v: unknown) => {
  const n = Number(String(v ?? "").replace(/[^\d.]/g, ""));
  return Number.isFinite(n) ? n : 0;
};

const normText = (v: unknown) => String(v ?? "").trim();

/** 채널 엑셀 양식 변경 대비: 여러 컬럼명 중 비어 있지 않은 첫 값 */
const pickRow = (row: Record<string, unknown>, ...keys: string[]): unknown => {
  for (const k of keys) {
    const v = row[k];
    if (v != null && String(v).trim() !== "") return v;
  }
  return keys.length ? row[keys[0]!] : undefined;
};

/** 토스 엑셀 안내 행(수정 불가/수정 가능) — 통합발주서에서 제외 */
const isTossTemplateHintText = (v: unknown) => {
  const s = String(v ?? "").trim();
  return s === "수정 불가" || s === "수정 가능";
};

const blankIfSameAsAddress = (address: string, message?: string) => {
  const a = normText(address);
  const m = normText(message);
  if (a && m && a === m) return ""; // 완전 동일하면 빈값
  return m;
};

// ExcelJS 셀 값이 object(하이퍼링크/리치텍스트/수식 등)인 경우를 문자열로 안전 변환
const cellText = (v: unknown) => {
  if (v == null) return "";

  // 기본 타입
  if (
    typeof v === "string" ||
    typeof v === "number" ||
    typeof v === "boolean"
  ) {
    return String(v).trim();
  }

  // ExcelJS 수식 셀: { formula, result }
  if (typeof v === "object" && v && "result" in (v as any)) {
    return String((v as any).result ?? "").trim();
  }

  // ExcelJS 하이퍼링크 셀: { text, hyperlink }
  if (typeof v === "object" && v && "text" in (v as any)) {
    return String((v as any).text ?? "").trim();
  }

  // ExcelJS 리치텍스트: { richText: [{ text: "..." }, ...] }
  if (typeof v === "object" && v && "richText" in (v as any)) {
    const parts = Array.isArray((v as any).richText) ? (v as any).richText : [];
    return parts
      .map((p: any) => p?.text ?? "")
      .join("")
      .trim();
  }

  // 최후: JSON으로 시도(디버깅에도 도움)
  try {
    return JSON.stringify(v);
  } catch {
    return String(v).trim();
  }
};

const pad2 = (n: number) => String(n).padStart(2, "0");

const formatKstDateTime = (v: unknown) => {
  if (!v) return "";

  // 1) Date 객체 (네이버에서 자주 나옴)
  if (v instanceof Date) {
    const yyyy = v.getFullYear();
    const mm = pad2(v.getMonth() + 1);
    const dd = pad2(v.getDate());
    const hh = pad2(v.getHours());
    const mi = pad2(v.getMinutes());
    return `${yyyy}-${mm}-${dd} ${hh}:${mi}`;
  }

  // 2) 문자열 정규화
  const s = String(v).trim();
  if (!s) return "";

  // 이미 문자열인 경우
  //  - 2026/01/31 08:42:13
  //  - 2026-01-31 08:42
  //  - 2026-01-31 08:42:13
  // 전부 → 2026-01-31 08:42

  // 슬래시 → 하이픈
  const withDash = s.replace(/\//g, "-");

  // 초 제거
  return withDash.replace(
    /^(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2})(:\d{2})?$/,
    (_, d, hm) => `${d} ${hm}`,
  );
};

const adminFields = (
  buyerName: string,
  buyerPhone: string,
  receiverName: string,
) => {
  const b = normName(buyerName);
  const r = normName(receiverName);
  if (b && r && b === r) {
    return {
      adminBuyerName: "기타",
      adminBuyerPhone: "010-6837-4121",
    };
  }
  return {
    adminBuyerName: `${b}/기타`,
    adminBuyerPhone: normPhone(buyerPhone),
  };
};

/**
 * row: 헤더 기반으로 뽑은 "한 행의 객체"
 * orderKeyFieldName: 각 채널 주문번호 컬럼명
 */
export const normalizeRow = (
  channel: TChannel,
  row: Record<string, unknown>,
): TStandardRow => {
  if (channel === "NAVER") {
    const orderKey = String(row["상품주문번호"] ?? "").trim();
    const buyerName = normName(row["구매자명"]);
    const buyerPhone = normPhone(row["구매자연락처"]);
    const receiverName = normName(row["수취인명"]);
    const address = normText(row["통합배송지"]);
    const message = blankIfSameAsAddress(address, row["배송메세지"] as string);

    const a = adminFields(buyerName, buyerPhone, receiverName);

    return {
      channel,
      orderKey,
      orderedAt: formatKstDateTime(row["주문일시"]),
      productName: String(row["상품명"] ?? "").trim(),
      quantity: toNumber(row["수량"]),
      receiverName,
      receiverPhone: normPhone(row["수취인연락처1"]),
      zipCode: String(row["우편번호"] ?? "").trim(),
      address,
      message,
      buyerName,
      buyerPhone,
      ...a,
      trackingNumber: "", // 통합단계에서는 비움
    };
  }

  if (channel === "TOSS") {
    let orderKey = String(row["주문상품번호"] ?? "").trim();
    if (
      isTossTemplateHintText(row["주문상품번호"]) ||
      isTossTemplateHintText(row["주문번호"])
    ) {
      orderKey = "";
    }
    const buyerName = normName(row["구매자명"]);
    const buyerPhone = normPhone(row["구매자 연락처"]);
    const receiverName = normName(row["수령인명"]);
    const product = normText(row["상품명"]);
    const option = normText(pickRow(row, "옵션명", "옵션"));
    const productName = option ? `${product} ${option}` : product;

    const a = adminFields(buyerName, buyerPhone, receiverName);

    return {
      channel,
      orderKey,
      orderedAt: formatKstDateTime(
        pickRow(row, "주문일시", "주문일자"),
      ),
      productName,
      quantity: toNumber(pickRow(row, "주문건수", "수량")),
      receiverName,
      receiverPhone: normPhone(row["수령인 연락처"]),
      zipCode: String(row["우편번호"] ?? "").trim(),
      address: normText(pickRow(row, "배송지", "주소")),
      message: normText(pickRow(row, "주문요청사항", "요청사항")),
      buyerName,
      buyerPhone,
      ...a,
      trackingNumber: "",
    };
  }

  if (channel === "COUPANG") {
    const orderKey = String(row["주문번호"] ?? "").trim();
    const buyerName = normName(row["구매자"]);
    const buyerPhone = normPhone(row["구매자전화번호"]);
    const receiverName = normName(row["수취인이름"]);

    const a = adminFields(buyerName, buyerPhone, receiverName);

    const productName =
      String(row["노출상품명(옵션명)"] ?? "").trim() ||
      String(row["등록상품명"] ?? "").trim();

    return {
      channel,
      orderKey,
      orderedAt: formatKstDateTime(row["주문일"]),
      productName,
      quantity: toNumber(row["구매수(수량)"]),
      receiverName,
      receiverPhone: normPhone(row["수취인전화번호"]),
      zipCode: String(row["우편번호"] ?? "").trim(),
      address: String(row["수취인 주소"] ?? "").trim(),
      message: String(row["배송메세지"] ?? "").trim(),
      buyerName,
      buyerPhone,
      ...a,
      trackingNumber: "",
    };
  }

  // MANDARINSPOON
  const orderKey = String(row["고객주문번호"] ?? "").trim();
  const buyerName = normName(row["보내는분성명"]);
  const buyerPhone = normPhone(row["보내는분전화번호"]);
  const receiverName = normName(row["받는분성명"]);

  const a = adminFields(buyerName, buyerPhone, receiverName);

  return {
    channel,
    orderKey,
    orderedAt: formatKstDateTime(row["주문일시"]),
    productName: cellText(row["품목명"]), // ✅ 변경
    quantity: toNumber(row["박스수량"]),
    receiverName,
    receiverPhone: normPhone(row["받는분전화번호"]),
    zipCode: cellText(row["받는분우편번호"]),
    address: cellText(row["받는분주소(전체, 분할)"]), // ✅ 변경
    message: cellText(row["배송메세지1"]), // ✅ 변경
    buyerName,
    buyerPhone,
    ...a,
    trackingNumber: "",
  };
};
