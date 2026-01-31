import JSZip from "jszip";
import { saveAs } from "file-saver";

export const arrayBufferToBlob = (ab: ArrayBuffer) =>
  new Blob([ab], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

export const downloadXlsx = (ab: ArrayBuffer, filename: string) => {
  saveAs(arrayBufferToBlob(ab), filename);
};

export const downloadZip = async (
  zipName: string,
  files: { filename: string; data: ArrayBuffer }[],
) => {
  const zip = new JSZip();
  files.forEach((f) => {
    zip.file(f.filename, f.data);
  });
  const blob = await zip.generateAsync({ type: "blob" });
  saveAs(blob, zipName);
};
