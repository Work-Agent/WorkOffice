export interface PluginApi {
  commands: {
    "pptx.guide": { input: Record<string, unknown>; output: unknown };
    "pptx.read": { input: Record<string, unknown>; output: unknown };
    "pptx.check": { input: Record<string, unknown>; output: unknown };
    "pptx.create": { input: Record<string, unknown>; output: unknown };
    "pptx.apply": { input: Record<string, unknown>; output: unknown };
    "pptx.load": { input: Record<string, unknown>; output: unknown };
    "pptx.save": { input: Record<string, unknown>; output: unknown };
    "pptx.saveAs": { input: Record<string, unknown>; output: unknown };
    "pptx.attach": { input: Record<string, unknown>; output: unknown };
    "pptx.pending": { input: Record<string, unknown>; output: unknown };
    "pptx.report": { input: Record<string, unknown>; output: unknown };
    "pptx.detach": { input: Record<string, unknown>; output: unknown };
    "pptx.ack": { input: Record<string, unknown>; output: unknown };
  };
}
