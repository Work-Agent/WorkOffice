// ─── office-docx 内部类型 ──────────────────────────────────────────────
//
// 这里只定义 DOCX 自己的计划形状与命令词汇（契约 §4.0：外壳统一、骨架不统一）。
// 宿主契约类型（CommitRequest / Receipt / RuntimeReport / ToolDecl）不在这里
// 重新定义——宿主是唯一事实源（WorkStudio src/shared/plugin-contract.ts）；
// 本包只按同样的结构产出，避免两份定义漂移。

/** body 顶层块种类；`sectPr` 为尾部隐藏块（不参与编辑与投影）。 */
export type BlockKind = 'p' | 'tbl' | 'sdt' | 'other' | 'sectPr';

/** 页面 / 执行点共用的读侧模型块。 */
export interface ModelBlock {
  /** 会话内稳定 id（页面历史 / 快照用；执行点按 `docxIndex` 生成）。 */
  id: string;
  /**
   * 基线位置身份（body 顶层下标）。
   *
   * 页面内新插入、尚未提交的块用**负数**占位（-1, -2 …）：这样命令坐标
   * 始终是同一个字段，不需要第二套「可见块序号」；提交后的 manifest 会把它
   * 换成真实下标（负数永远不可能命中基线切片，因此 plan 里必然写成 xml）。
   */
  docxIndex: number | null;
  kind: BlockKind;
  /** 顶层 XML（段落 / 表格 / sdt …）。 */
  xml: string;
  /** 纯文本投影（表格为空串）。 */
  text: string;
  /** 0 = 正文，1..6 = Heading N。 */
  heading: number;
  /** 尾部 `w:sectPr` 等不可见块：不渲染、不可编辑、投影里跳过。 */
  hidden?: boolean;
}

/** SavePlan 的一项：原样回写基线切片，或写入新的顶层 XML。 */
export type SaveBlock =
  | { kind: 'original'; docxIndex: number }
  | { kind: 'xml'; xml: string; docxIndex?: number | null };

/**
 * DOCX 的格式内部计划（契约 §9.5 的 `plan`）。
 *
 * `saveBlocks` 是 `word/document.xml` 顶层子节点的**最终序列**——这是 DOCX 的
 * 文件格式性质，不是跨格式抽象（契约 I5）。未改块用 `{kind:'original'}` 指回
 * 基线切片，`exec` 只重写 `word/document.xml` 一个 zip 条目。
 */
export interface SavePlan {
  kind: 'save' | 'check' | 'reject';
  saveBlocks?: SaveBlock[];
  /** `docxIndex → saveBlocks 下标`（诊断 / 逆计划用）。 */
  saveBlockIndexByDocx?: Record<number, number>;
  changedCount: number;
  deletedCount: number;
  totalOriginal: number;
  /** 本批成功应用的命令条数（回执 `applied` 的取值）。 */
  applied?: number;
  /**
   * 工具参数里的 `dryRun` 由 plan 带出（宿主当前不把工具入参的 dryRun
   * 透传给 exec）；exec 见 true 即只校验、不产字节。
   */
  dryRun?: boolean;
  /** 尾部 sectPr 被改写（正常应为 false：它原样保留）。 */
  trailingSectPrChange: boolean;
  /** `check` 的数据或拒绝信息。 */
  audit?: unknown;
  rejected?: Rejection[];
  warnings?: string[];
}

export interface Rejection {
  index: number;
  code: string;
  message: string;
}

/** 一次提交里被拒的第一条命令：批级原子，整批不落盘（`isolation: 'batch'`）。 */
export class CommandRejected extends Error {
  constructor(
    readonly index: number,
    readonly code: string,
    message: string,
  ) {
    super(message);
    this.name = 'CommandRejected';
  }
}

// ─── 命令词汇（catalog 即开关：未实现的不进目录） ─────────────────────

/** 跨块选区（按可见序归一化，偏移是段内字符偏移）。 */
export interface BlockRange {
  /** 起点块 `docxIndex`。 */
  from: number;
  fromOffset: number;
  /** 终点块 `docxIndex`。 */
  to: number;
  toOffset: number;
}

export interface CommandBase {
  /** 目标块的 `docxIndex`（= sidecar 打印的索引 = 命令目标）。 */
  index?: number;
}

export type DocxCommand =
  | { op: 'insertParagraph'; after?: number | null; text?: string; heading?: number }
  | { op: 'setText'; index: number; text: string }
  | { op: 'setHeadingLevel'; index: number; level: number }
  | {
      op: 'setTextStyle';
      index: number;
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      /** 字号（磅）。 */
      size?: number;
      /** `RRGGBB`。 */
      color?: string;
      align?: 'left' | 'center' | 'right' | 'both';
    }
  | { op: 'deleteBlock'; index: number }
  /** `to` = 目标块之后（null = 文档开头）。 */
  | { op: 'moveBlock'; index: number; to: number | null }
  | { op: 'findReplace'; find: string; replace: string; all?: boolean }
  | { op: 'check' };

/** 提交后页面应用的 baseline sync 清单（契约 §8.3：docx 为 manifest）。 */
export interface SyncManifestBlock {
  docxIndex: number;
  /** sha256 前 16 字节的十六进制。 */
  hash: string;
  /** 页面没有这段切片时回传（新增 / 被重写的块）。 */
  xml?: string;
}

export interface SyncManifest {
  blocks: SyncManifestBlock[];
  hidden?: Array<{ docxIndex: number; kind: 'sectPr' | 'sdtClose' }>;
}

/** 页面代理回执里回给模型的投影。 */
export interface BlockProjection {
  kind: 'docx';
  ref?: string;
  version: string;
  epoch: number;
  docRev: string;
  blocks: Array<{ index: number; type: string; text: string }>;
  text: string;
}
