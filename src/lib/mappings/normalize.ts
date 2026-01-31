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

const blankIfSameAsAddress = (address: string, message?: string) => {
  const a = normText(address);
  const m = normText(message);
  if (a && m && a === m) return ""; // 완전 동일하면 빈값
  return m;
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
      adminBuyerName: "귤수저",
      adminBuyerPhone: "010-6837-4121",
    };
  }
  return {
    adminBuyerName: `${b}/귤수저`,
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
      orderedAt: String(row["주문일시"] ?? "").trim(),
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
    const orderKey = String(row["주문상품번호"] ?? "").trim();
    const buyerName = normName(row["구매자명"]);
    const buyerPhone = normPhone(row["구매자 연락처"]);
    const receiverName = normName(row["수령인명"]);
    const product = normText(row["상품명"]);
    const option = normText(row["옵션"]); // ⚠️ 토스 파일의 옵션 컬럼명이 정확히 "옵션"인지 확인 필요
    const productName = option ? `${product} ${option}` : product;

    const a = adminFields(buyerName, buyerPhone, receiverName);

    return {
      channel,
      orderKey,
      orderedAt: String(row["주문일자"] ?? "").trim(),
      productName,
      quantity: toNumber(row["수량"]),
      receiverName,
      receiverPhone: normPhone(row["수령인 연락처"]),
      zipCode: String(row["우편번호"] ?? "").trim(),
      address: String(row["주소"] ?? "").trim(),
      message: String(row["요청사항"] ?? "").trim(),
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
      orderedAt: String(row["주문일"] ?? "").trim(),
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
    orderedAt: String(row["주문일시"] ?? "").trim(), // 없으면 빈값
    productName: String(row["품목명"] ?? "").trim(),
    quantity: toNumber(row["박스수량"]),
    receiverName,
    receiverPhone: normPhone(row["받는분전화번호"]),
    zipCode: String(row["받는분우편번호"] ?? "").trim(),
    address: String(row["받는분주소"] ?? "").trim(),
    message: String(row["배송메세지1"] ?? "").trim(),
    buyerName,
    buyerPhone,
    ...a,
    trackingNumber: "",
  };
};
