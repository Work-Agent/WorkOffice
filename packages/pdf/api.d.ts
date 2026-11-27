export interface PluginApi {
  commands: {
    "pdf.read": { input: Record<string, unknown>; output: unknown };
    "pdf.load": { input: Record<string, unknown>; output: unknown };
    "pdf.chunk": { input: Record<string, unknown>; output: unknown };
    "pdf.saveCopy": { input: Record<string, unknown>; output: unknown };
  };
}
