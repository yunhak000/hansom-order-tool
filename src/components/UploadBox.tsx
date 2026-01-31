"use client";

import { useCallback } from "react";
import { useDropzone } from "react-dropzone";

type Props = {
  title: string;
  description?: string;
  accept?: Record<string, string[]>;
  multiple?: boolean;
  onFiles: (files: File[]) => void;
};

export const UploadBox = ({
  title,
  description,
  accept = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [
      ".xlsx",
    ],
  },
  multiple = true,
  onFiles,
}: Props) => {
  const onDrop = useCallback(
    (acceptedFiles: File[]) => {
      if (!acceptedFiles?.length) return;
      onFiles(acceptedFiles);
    },
    [onFiles],
  );

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept,
    multiple,
  });

  return (
    <div className="rounded-2xl border border-neutral-200 bg-white p-5 shadow-sm">
      <div className="flex items-start justify-between gap-3">
        <div>
          <div className="text-base font-semibold">{title}</div>
          {description && (
            <div className="mt-1 text-sm text-neutral-500">{description}</div>
          )}
        </div>
        <div className="text-xs text-neutral-500">.xlsx</div>
      </div>

      <div
        {...getRootProps()}
        className={[
          "mt-4 cursor-pointer rounded-xl border-2 border-dashed p-6 text-center transition",
          isDragActive
            ? "border-neutral-500 bg-neutral-50"
            : "border-neutral-200 hover:bg-neutral-50",
        ].join(" ")}
      >
        <input {...getInputProps()} />
        <div className="text-sm font-medium">
          {isDragActive
            ? "여기에 놓아주세요"
            : "파일을 드래그하거나 클릭해서 업로드"}
        </div>
        <div className="mt-1 text-xs text-neutral-500">
          여러 개 파일을 한 번에 올려도 자동으로 네이버/토스/쿠팡/귤수저로
          분류합니다.
        </div>
      </div>
    </div>
  );
};
