"use client";

import type { WarehouseData } from "@/lib/local/types";

export const sheetEndpointKey = "tile_warehouse_sheet_endpoint";
export const sheetAppKeyKey = "tile_warehouse_sheet_app_key";

export type SheetConfig = {
  endpoint: string;
  appKey: string;
};

type SheetResponse = {
  ok: boolean;
  version?: number;
  revision?: string;
  data?: WarehouseData;
  error?: string;
};

export type SheetSnapshotResult = {
  data: WarehouseData;
  revision: string;
};

export type SheetMeta = {
  revision: string;
};

let callbackCounter = 0;
const metaTimeoutMs = 30000;
const snapshotTimeoutMs = 60000;
const mutateTimeoutMs = 90000;

function getWindow() {
  return typeof window === "undefined" ? null : window;
}

export function getStoredSheetConfig(): SheetConfig {
  const win = getWindow();

  if (!win) {
    return getBuildSheetConfig();
  }

  const storedEndpoint = win.localStorage.getItem(sheetEndpointKey) ?? "";
  const storedAppKey = win.localStorage.getItem(sheetAppKeyKey) ?? "";
  const buildConfig = getBuildSheetConfig();

  return {
    endpoint: storedEndpoint || buildConfig.endpoint,
    appKey: storedAppKey || buildConfig.appKey,
  };
}

export function saveStoredSheetConfig(config: SheetConfig) {
  const win = getWindow();

  if (!win) {
    return;
  }

  win.localStorage.setItem(sheetEndpointKey, config.endpoint.trim());
  win.localStorage.setItem(sheetAppKeyKey, config.appKey.trim());
}

export function isSheetConfigured(config = getStoredSheetConfig()) {
  return Boolean(config.endpoint.trim());
}

function getBuildSheetConfig(): SheetConfig {
  return {
    endpoint: process.env.NEXT_PUBLIC_GOOGLE_SCRIPT_URL ?? "",
    appKey: process.env.NEXT_PUBLIC_GOOGLE_APP_KEY ?? "",
  };
}

function getConfigFileUrl() {
  const basePath = process.env.NEXT_PUBLIC_BASE_PATH ?? "";
  return `${basePath}/google-sheet-config.json`;
}

export async function loadSheetConfig() {
  const buildConfig = getBuildSheetConfig();

  if (isSheetConfigured(buildConfig)) {
    return buildConfig;
  }

  try {
    const response = await fetch(getConfigFileUrl(), { cache: "no-store" });

    if (!response.ok) {
      return getStoredSheetConfig();
    }

    const fileConfig = (await response.json()) as Partial<SheetConfig>;
    const resolvedConfig = {
      endpoint: String(fileConfig.endpoint ?? "").trim(),
      appKey: String(fileConfig.appKey ?? "").trim(),
    };

    if (isSheetConfigured(resolvedConfig)) {
      saveStoredSheetConfig(resolvedConfig);
      return resolvedConfig;
    }
  } catch {
    return getStoredSheetConfig();
  }

  return getStoredSheetConfig();
}

function jsonpRequest(endpoint: string, params: Record<string, string>, timeoutMs = snapshotTimeoutMs) {
  const win = getWindow();

  if (!win) {
    return Promise.reject(new Error("Khong chay trong trinh duyet."));
  }

  return new Promise<SheetResponse>((resolve, reject) => {
    const callbackName = `__tileWarehouseSheetCallback${Date.now()}${callbackCounter++}`;
    const url = new URL(endpoint);
    const script = document.createElement("script");
    let settled = false;

    Object.entries(params).forEach(([key, value]) => {
      url.searchParams.set(key, value);
    });
    url.searchParams.set("callback", callbackName);

    function cleanup() {
      if (timeoutId) {
        window.clearTimeout(timeoutId);
      }

      delete (window as unknown as Record<string, unknown>)[callbackName];
      script.remove();
    }

    (window as unknown as Record<string, (response: SheetResponse) => void>)[callbackName] = (
      response,
    ) => {
      if (settled) {
        return;
      }

      settled = true;
      cleanup();
      resolve(response);
    };

    script.onerror = () => {
      if (settled) {
        return;
      }

      settled = true;
      cleanup();
      reject(new Error("Khong ket noi duoc Google Apps Script."));
    };

    const timeoutId = window.setTimeout(() => {
      if (settled) {
        return;
      }

      settled = true;
      cleanup();
      reject(new Error("Google Sheet phản hồi quá lâu. Có thể Sheet hoặc Apps Script đang chậm, hãy thử lại sau vài giây."));
    }, timeoutMs);

    script.src = url.toString();
    document.body.appendChild(script);
  });
}

function getRevision(response: SheetResponse) {
  return String(response.revision ?? "");
}

async function assertSnapshotOk(response: SheetResponse): Promise<SheetSnapshotResult> {
  if (!response.ok || !response.data) {
    throw new Error(response.error || "Google Sheet tra ve du lieu khong hop le.");
  }

  return {
    data: response.data,
    revision: getRevision(response),
  };
}

async function assertMetaOk(response: SheetResponse): Promise<SheetMeta> {
  if (!response.ok) {
    throw new Error(response.error || "Khong kiem tra duoc thay doi Google Sheet.");
  }

  const revision = getRevision(response);

  if (!revision) {
    throw new Error("Apps Script chua ho tro kiem tra thay doi.");
  }

  return { revision };
}

export async function fetchSheetMeta(config = getStoredSheetConfig(), timeoutMs = metaTimeoutMs) {
  const response = await jsonpRequest(config.endpoint, {
    action: "meta",
    key: config.appKey,
  }, timeoutMs);
  return assertMetaOk(response);
}

export async function fetchSheetSnapshotWithMeta(config = getStoredSheetConfig(), timeoutMs = snapshotTimeoutMs) {
  const response = await jsonpRequest(config.endpoint, {
    action: "snapshot",
    key: config.appKey,
  }, timeoutMs);
  return assertSnapshotOk(response);
}

export async function fetchSheetSnapshot(config = getStoredSheetConfig()) {
  return (await fetchSheetSnapshotWithMeta(config)).data;
}

export async function mutateSheetWithMeta(
  type: string,
  payload: Record<string, unknown>,
  config = getStoredSheetConfig(),
) {
  const response = await jsonpRequest(config.endpoint, {
    action: "mutate",
    type,
    key: config.appKey,
    payload: JSON.stringify(payload),
  }, mutateTimeoutMs);
  return assertSnapshotOk(response);
}

function wait(ms: number) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

export async function postMutateSheetWithMeta(
  type: string,
  payload: Record<string, unknown>,
  config = getStoredSheetConfig(),
) {
  const before = await fetchSheetMeta(config).catch(() => null);
  const body = new URLSearchParams();

  body.set("action", "mutate");
  body.set("type", type);
  body.set("key", config.appKey);
  body.set("payload", JSON.stringify(payload));

  await fetch(config.endpoint, {
    method: "POST",
    mode: "no-cors",
    body,
  });

  for (let attempt = 0; attempt < 20; attempt += 1) {
    await wait(1000);

    const meta = await fetchSheetMeta(config, 10000).catch(() => null);

    if (!before || (meta && meta.revision !== before.revision)) {
      return fetchSheetSnapshotWithMeta(config, snapshotTimeoutMs);
    }
  }

  return fetchSheetSnapshotWithMeta(config, snapshotTimeoutMs);
}

export async function mutateSheet(
  type: string,
  payload: Record<string, unknown>,
  config = getStoredSheetConfig(),
) {
  return (await mutateSheetWithMeta(type, payload, config)).data;
}

export function formDataToPayload(formData: FormData) {
  const payload: Record<string, FormDataEntryValue | FormDataEntryValue[]> = {};

  formData.forEach((value, key) => {
    const current = payload[key];

    if (Array.isArray(current)) {
      current.push(value);
    } else if (current !== undefined) {
      payload[key] = [current, value];
    } else {
      payload[key] = value;
    }
  });

  return payload;
}
