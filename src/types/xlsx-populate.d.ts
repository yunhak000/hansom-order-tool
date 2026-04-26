declare module "xlsx-populate" {
  export type XlsxPopulateWorkbook = {
    outputAsync: (opts: { type: "nodebuffer" }) => Promise<Buffer>;
  };

  type XlsxPopulateOptions = {
    password?: string;
  };

  const XlsxPopulate: {
    fromDataAsync: (
      data: Buffer | Uint8Array | ArrayBuffer,
      options?: XlsxPopulateOptions,
    ) => Promise<XlsxPopulateWorkbook>;
  };

  export default XlsxPopulate;
}
