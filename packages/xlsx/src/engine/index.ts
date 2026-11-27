/**
 * xlsx 引擎入口：definePlugin（WorkStudio 0.3 插件框架）。
 *
 * 命令面只有两组：
 *   agent   guide / read        —— `expose.agent: true`，常开
 *   page    page                —— 页面通道，`agent: false` + 租约令牌
 * 写命令（create / apply）在 P2 / P3 接入；页面通道的 op 会随之扩展。
 */

import { definePlugin } from '@work-studio/sdk'

import { apply } from './commands/apply'
import { createWorkbook } from './commands/create'
import { guide } from './commands/guide'
import { page } from './commands/page'
import { read } from './commands/read'
import { resetHashes } from './hash'
import { resetSessions } from './sessions'
import { resetWasmEngine } from './wasm'
import { resetViews } from './read'
import { xlsxResource } from './resource'
import { clearScope, setScope } from './scope'
import { bindNotifier, resetLeases } from './state'

const REF = { type: 'string', required: true } as const

export const workstudio = definePlugin({
  events: true,
  resources: {
    xlsx: xlsxResource,
  },
  commands: {
    'xlsx.guide': {
      summary:
        '取 .xlsx 的命令词汇：无参 → 四组 op 名与数量；group → 组内 schema；op → 单条 schema 与示例。' +
        '写命令前先调它，不要凭记忆写 op。',
      args: {
        ref: REF,
        group: { enum: ['cells', 'structure', 'sheets', 'objects'] },
        op: { type: 'string' },
      },
      expose: { agent: true },
      run: (input) => guide(input as Record<string, unknown>),
    },

    'xlsx.read': {
      summary:
        '读单个工作簿：overview（context）、工作表清单（sheets）、区域值（range）、' +
        '公式（formulas）、查找（find）、包结构诊断（audit）。只读，不写盘；' +
        '结果首行带 hash/size 令牌，写命令用它做基线。',
      args: {
        ref: REF,
        what: { enum: ['context', 'sheets', 'range', 'formulas', 'find', 'audit'] },
        sheet: { type: 'string' },
        range: { type: 'string' },
        start: { type: 'string' },
        end: { type: 'string' },
        query: { type: 'string' },
        regex: { type: 'boolean' },
        matchCase: { type: 'boolean' },
        limit: { type: 'number' },
      },
      expose: { agent: true },
      run: (input) => read(input as Record<string, unknown>),
    },

    'xlsx.create': {
      summary:
        '从零新建 .xlsx：ref 是尚不存在的目标资源（工作区内、父目录已存在）。' +
        'sheets 是 [{ name, rows }]，rows 二维数组里以 = 开头写成公式；也支持 rows / csv 建单表。' +
        '空内容 = 空白工作簿。目标已存在时拒绝，改已有工作簿请用 xlsx.apply。',
      args: {
        ref: REF,
        // sheets / rows / csv：数组或 JSON 字符串都接受，所以不声明（同 apply）。
      },
      expose: { agent: true },
      run: (input, ctx) => createWorkbook(input as Record<string, unknown>, ctx),
    },

    'xlsx.apply': {
      summary:
        '按顺序执行一批命令（上限 200，op 见 xlsx.guide）。' +
        'cells / structure / sheets 组不需要页面；objects 组（图表/透视表/条件格式/批注/筛选等）需要打开页面，' +
        '否则整批回 command.page-unavailable。' +
        '有页面时命令进页面事务：含 A1 坐标的结构类 op 需 structureEpoch，读内容型 op（fill/copy/sort/find）需 docRev。' +
        '被拒时回执给出 index 与 code。禁止用 write/edit 改 .xlsx。',
      args: {
        ref: REF,
        baseHash: { type: 'string' },
        docRev: { type: 'number' },
        structureEpoch: { type: 'number' },
        dryRun: { type: 'boolean' },
        // `commands` 不声明：要同时接受对象数组与 JSON 字符串（宿主只校验声明项）。
      },
      expose: { agent: true },
      run: (input, ctx) => apply(input as Record<string, unknown>, ctx),
    },

    'xlsx.page': {
      summary:
        '页面通道（内部）：status/open/chunk/meta/range 只读；attach 发放令牌；' +
        'report/detach 需令牌。助手不要调用。',
      args: {
        ref: REF,
        pageId: { type: 'string' },
        token: { type: 'string' },
        op: {
          enum: [
            'status',
            'open',
            'chunk',
            'meta',
            'range',
            'save',
            'saveAs',
            'saveBegin',
            'saveChunk',
            'saveCommit',
            'attach',
            'pending',
            'ack',
            'report',
            'detach',
            'draft',
            'draftStatus',
            'draftLoad',
            'draftClear',
          ],
        },
        // `input` 故意不声明：宿主对声明过的参数做 `typeof` 严格校验，
        // 而这里要传对象甚至 Uint8Array（save/chunk）。未声明的字段原样透传。
      },
      expose: { agent: false },
      run: (input, ctx) => page(input as Record<string, unknown>, ctx),
    },
  },
  lifecycle: {
    activate(scope) {
      setScope(scope)
      // agent → 页面 的请求只推 id/type；正文由页面用 `pending` 取。
      bindNotifier((_ref, request) => {
        scope.events.emit('request', { id: request.id, type: request.type })
      })
    },
    deactivate() {
      resetLeases()
      resetViews()
      resetHashes()
      resetWasmEngine()
      resetSessions()
      clearScope()
    },
  },
})

export default workstudio
