export interface PluginApi {
  commands: {
    "docx-pro.guide": { input: Record<string, unknown>; output: unknown };
    "docx-pro.read": { input: Record<string, unknown>; output: unknown };
    "docx-pro.apply": { input: Record<string, unknown>; output: unknown };
    "docx-pro.create": { input: Record<string, unknown>; output: unknown };
    "docx-pro.load": { input: Record<string, unknown>; output: unknown };
    "docx-pro.save": { input: Record<string, unknown>; output: unknown };
    "docx-pro.attach": { input: Record<string, unknown>; output: unknown };
    "docx-pro.pending": { input: Record<string, unknown>; output: unknown };
    "docx-pro.report": { input: Record<string, unknown>; output: unknown };
    "docx-pro.detach": { input: Record<string, unknown>; output: unknown };
    "docx-pro.ack": { input: Record<string, unknown>; output: unknown };
  };
}
