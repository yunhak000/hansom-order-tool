import localforage from "localforage";
import { TAppState } from "@/lib/types";

const KEY = "hansom-order-tool:v1";

localforage.config({
  name: "hansom-order-tool",
  storeName: "app",
});

export const saveState = async (state: TAppState) => {
  // ArrayBuffer는 localforage가 저장 가능
  await localforage.setItem(KEY, state);
};

export const loadState = async (): Promise<TAppState | null> => {
  const v = await localforage.getItem<TAppState>(KEY);
  return v ?? null;
};

export const clearState = async () => {
  await localforage.removeItem(KEY);
};
