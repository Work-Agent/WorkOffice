// ─── XLSX codec（WorkStudio 自有） ───────────────────────────────────
//
// codec 不依赖 GenOffice：用 fflate + fast-xml-parser 读 OOXML 包，
// 生成 sheet/cell 快照供 sidecar 与 presenter 使用。页面只消费快照。

import { registerCodec, registerPresenter, registerSidecar, registerStore } from './sdk'
import { MIME } from './ura-types'
import { cellText, nameOf, parseWorkbook, pick, workbookText, type XlsxDoc } from './shared/xlsx-parse'
import type {
  Codec,
  Patch,
  Presenter,
  RawContent,
  SidecarProvider,
  Store,
} from './ura-types'

const CODEC_ID = 'xlsx-pro'

const codec: Codec = {
  id: CODEC_ID,
  keys: [{ transport: '*', mime: MIME.xlsx }],
  extensions: ['xlsx', 'xlsm', 'xltx'],
  patchType: 'bytes',

  decode(raw: RawContent, ref: string) {
    const sheets = parseWorkbook(raw.bytes)
    return {
      doc: { name: nameOf(ref), sheets } satisfies XlsxDoc,
      // workbookText 同时是 agent 文本与 grep 锚点。
      anchorMap: workbookText(sheets),
    }
  },

  locate(_decoded, anchor) {
    return { start: anchor.start, end: anchor.end }
  },

  // 字节只由 exec 产生（契约 I8 / I9）；整份写通道对本插件关闭。
  encode(): Patch {
    throw new Error('command.forbidden')
  },
}

const agentSidecar: SidecarProvider = {
  key: { transport: '*', mime: MIME.xlsx, codec: CODEC_ID },
  sidecar(decoded) {
    const text =
      typeof decoded.anchorMap === 'string'
        ? decoded.anchorMap
        : workbookText((decoded.doc as XlsxDoc).sheets)
    return {
      type: 'text',
      content: text || '(空 XLSX 工作簿)',
    }
  },
}

const editorPresenter: Presenter = {
  key: { transport: '*', mime: MIME.xlsx, codec: CODEC_ID },
  type: 'page',
  surface: 'xlsx-pro',
  requires: ['write'],
  present(decoded, raw, ref) {
    const doc = decoded.doc as XlsxDoc
    return {
      type: 'page',
      surface: 'xlsx-pro',
      page: '',
      // 字节经 `bytes` 分块通道取（content: 'pulled'）；data 只放页面启动
      // 需要的元数据与工作表快照（页面不用再解析一遍）。
      data: {
        name: doc.name,
        ref,
        size: raw.bytes.byteLength,
        version: raw.version,
        sheets: doc.sheets,
      },
    }
  },
}

const sheetStore: Store = {
  key: { transport: '*', mime: MIME.xlsx, codec: CODEC_ID },
  schema: {
    sheets: { name: 'string', path: 'string', cells: 'number' },
    cells: { sheet: 'string', address: 'string', value: 'any', formula: 'string' },
  },
  query(decoded, query) {
    const doc = decoded.doc as XlsxDoc
    const where = query.where ?? {}
    const rows = doc.sheets.flatMap((sheet) =>
      Object.entries(sheet.cells).map(([address, cell]) => ({
        sheet: sheet.name,
        address,
        value: cell.value,
        formula: cell.formula ?? '',
      })),
    )
    const filtered =
      typeof where['sheet'] === 'string'
        ? rows.filter((row) => row.sheet === where['sheet'])
        : rows
    const limit = Math.max(1, Math.min(2_000, Number(query.limit ?? 200)))
    return {
      total: filtered.length,
      rows: query.select?.length
        ? filtered.slice(0, limit).map((row) => pick(row, query.select!))
        : filtered.slice(0, limit),
    }
  },
}

registerCodec(codec)
registerSidecar(agentSidecar)
registerStore(sheetStore)
registerPresenter(editorPresenter)
