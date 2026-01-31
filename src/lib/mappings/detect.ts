import { TChannel } from "@/lib/types";

const hasAll = (headers: string[], required: string[]) =>
  required.every((r) => headers.includes(r));

export const detectChannelByHeaders = (headers: string[]): TChannel | null => {
  // 네이버
  if (hasAll(headers, ["상품주문번호", "수취인명", "송장번호"])) return "NAVER";

  // 토스
  if (hasAll(headers, ["주문상품번호", "수령인명", "송장번호"])) return "TOSS";

  // 쿠팡
  if (hasAll(headers, ["주문번호", "수취인이름", "운송장번호"]))
    return "COUPANG";

  // 귤수저(개인)
  if (hasAll(headers, ["고객주문번호", "보내는분성명", "받는분성명"]))
    return "MANDARINSPOON";

  return null;
};
