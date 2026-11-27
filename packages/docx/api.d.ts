export interface PluginApi {
  commands: {
    "docx.guide": { input: Record<string, unknown>; output: unknown };
    "docx.read": { input: Record<string, unknown>; output: unknown };
    "docx.apply": { input: Record<string, unknown>; output: unknown };
    "docx.create": { input: Record<string, unknown>; output: unknown };
    "docx.load": { input: Record<string, unknown>; output: unknown };
    "docx.save": { input: Record<string, unknown>; output: unknown };
    "docx.attach": { input: Record<string, unknown>; output: unknown };
    "docx.pending": { input: Record<string, unknown>; output: unknown };
    "docx.report": { input: Record<string, unknown>; output: unknown };
    "docx.detach": { input: Record<string, unknown>; output: unknown };
    "docx.ack": { input: Record<string, unknown>; output: unknown };
  };
}
