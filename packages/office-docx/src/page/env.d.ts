// 页面运行在 WorkStudio 的沙箱里（contextIsolation + sandbox，无 Node）。
// 只声明本页用到的宿主面（WorkStudio src/preload/extension.ts）。

export {};

export interface BusMessage {
  type: string;
  payload?: { meta?: { id?: string }; content?: string };
  id?: string;
  kind?: 'context' | 'apply';
  txnId?: string;
  version?: string;
  epoch?: number;
  commands?: unknown[];
  dryRun?: boolean;
  theme?: string;
  appearance?: string;
}

export interface Receipt {
  ref: string;
  txnId: string;
  version: string;
  epoch: number;
  sync?: {
    kind: 'none' | 'version' | 'manifest' | 'bytes';
    blocks?: unknown[];
    dataUrl?: string;
  };
  applied: number;
  rejected?: Array<{ index: number; code: string; message: string }>;
  warnings?: string[];
  notices?: string[];
  audit?: unknown;
}

export interface PageSessionInfo {
  id: string;
  ref: string;
  version: string;
  epoch: number;
  codec: string | null;
  surface: string;
  page: string;
  data: unknown;
  degraded: boolean;
  write: 'none' | 'bytes' | 'plan';
  tier?: 'normal' | 'lite' | 'readOnly' | 'refused';
  tools?: string[];
  runtime?: unknown;
  size?: number;
}

export interface ExtensionApi {
  ura: {
    page: {
      get<T = PageSessionInfo>(id: string): Promise<T>;
      commit<T = Receipt>(id: string, request: unknown): Promise<T>;
      bytes(id: string, window?: { start?: number; end?: number }): Promise<Uint8Array>;
      reload<T = PageSessionInfo>(id: string): Promise<T>;
      agentResult(id: string, message: unknown): Promise<unknown>;
    };
  };
  bus: {
    on(cb: (message: BusMessage) => void): () => void;
    send(message: Record<string, unknown>): void;
    loaded(): void;
  };
}

declare global {
  interface Window {
    extension: ExtensionApi;
  }
}
