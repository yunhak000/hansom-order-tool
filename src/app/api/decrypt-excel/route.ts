import { NextResponse } from "next/server";
import XlsxPopulate from "xlsx-populate";

const XLSX_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const DECRYPT_TIMEOUT_MS = 15000;

const withTimeout = async <T,>(promise: Promise<T>, timeoutMs: number) => {
  let timeoutId: ReturnType<typeof setTimeout> | null = null;
  const timeoutPromise = new Promise<never>((_, reject) => {
    timeoutId = setTimeout(() => {
      reject(new Error("DECRYPT_TIMEOUT"));
    }, timeoutMs);
  });

  try {
    return await Promise.race([promise, timeoutPromise]);
  } finally {
    if (timeoutId) clearTimeout(timeoutId);
  }
};

const isLikelyInvalidPassword = (err: unknown) => {
  const msg = String(err ?? "").toLowerCase();
  return (
    msg.includes("password") ||
    msg.includes("decrypt") ||
    msg.includes("unsupported encryption")
  );
};

export async function POST(req: Request) {
  const startAt = Date.now();

  try {
    const formData = await req.formData();
    const file = formData.get("file");
    const password = String(formData.get("password") ?? "");

    if (!(file instanceof File)) {
      return NextResponse.json(
        { code: "BAD_REQUEST", message: "업로드 파일을 찾지 못했어요." },
        { status: 400 },
      );
    }

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      return NextResponse.json(
        { code: "UNSUPPORTED_FILE", message: ".xlsx 파일만 업로드할 수 있어요." },
        { status: 400 },
      );
    }

    if (!password.trim()) {
      return NextResponse.json(
        { code: "PASSWORD_REQUIRED", message: "비밀번호를 입력해주세요." },
        { status: 400 },
      );
    }

    const input = Buffer.from(await file.arrayBuffer());

    const workbook = await withTimeout(
      XlsxPopulate.fromDataAsync(input, { password }),
      DECRYPT_TIMEOUT_MS,
    );

    const decrypted = await withTimeout(
      workbook.outputAsync({ type: "nodebuffer" }) as Promise<Buffer>,
      DECRYPT_TIMEOUT_MS,
    );

    console.info("[decrypt-excel] success", {
      filename: file.name,
      inputBytes: input.byteLength,
      outputBytes: decrypted.byteLength,
      elapsedMs: Date.now() - startAt,
    });

    return new NextResponse(decrypted, {
      status: 200,
      headers: {
        "Content-Type": XLSX_CONTENT_TYPE,
        "Cache-Control": "no-store",
      },
    });
  } catch (err) {
    const msg = String(err ?? "");
    const isTimeout = msg.includes("DECRYPT_TIMEOUT");
    const invalidPassword = isLikelyInvalidPassword(err);

    console.warn("[decrypt-excel] failed", {
      reason: isTimeout
        ? "timeout"
        : invalidPassword
          ? "invalid_password"
          : "unknown",
      detail: msg,
      elapsedMs: Date.now() - startAt,
    });

    if (isTimeout) {
      return NextResponse.json(
        {
          code: "DECRYPT_TIMEOUT",
          message:
            "복호화 처리 시간이 오래 걸리고 있어요. 잠시 후 다시 시도해주세요.",
        },
        { status: 408 },
      );
    }

    if (invalidPassword) {
      return NextResponse.json(
        { code: "INVALID_PASSWORD", message: "비밀번호가 올바르지 않아요." },
        { status: 401 },
      );
    }

    return NextResponse.json(
      {
        code: "DECRYPT_FAILED",
        message:
          "파일을 열지 못했어요. 비밀번호를 확인하거나 파일 형식을 다시 확인해주세요.",
      },
      { status: 400 },
    );
  }
}
