# docx-pro（Word · `.docx`）

把两个既有实现合成**一个插件**：

- `packages/docx` 的**完整编辑器**（GenOffice Docs fork：ProseMirror、分页、Ribbon、
  审阅、打印/PDF 导出）——整份源码复制进 `src/editor`，一个字节都没丢；插件之间
  不互相 import，`packages/docx` 可以独立存在、独立升级。
- `packages/office-docx` 的**计划型写协议**（SavePlan、块级 XML 切片、manifest
  sync、CAS、只读投影、命令目录）——无页面也能跑。

- 插件 id：`docx-pro`
- 清单：`apiVersion: 2` / `write: 'plan'` / `content: 'pulled'` / `model: 'none'`
- 工具：**3 个**（`docx_guide` / `docx_read` / `docx_apply`），目录里 **41 条命令**
  （blocks 8 · editor 23 · actions 10）——声明体积 ≈1.7 KB，而不是逐工具面的 32.5 KB
- 技能：`skills/docx/SKILL.md`

---

## 1. 一个插件、两种写来源、一个字节出口

```
有页面（人在场）
  编辑器 PM 事务（executeTool / apply_ops / docx_edit 映射）
    → 编辑器自己的 save()（重新解析基线、刷新 dirty）
    → serialize（buildDocBytes）
    → plan { kind: 'bytes' }        ← 页面交整份字节
                          ↘
                          exec（校验包结构 / 预算）→ 权威层 CAS → 落盘

无页面（后台 / 批处理）
  docx_edit 命令批
    → 执行点 plan（XML 切片，未改块逐字节不变）
    → plan { kind: 'save' }
                          ↘
                          exec（只重写 word/document.xml）→ 权威层 CAS → 落盘
```

`exec` 是契约里唯一的字节产生者；两条路径都经过它，版本、`epoch`、撤销与冲突
语义只有一份。

为什么必须 `write: 'plan'`：宿主把写形态做成**每插件单值**（`page-sessions.commit`
用同一个 `writeModeOf` 裁决 `bytes` 与 `plan`）。要同时保留「页面整份序列化」与
「无页面 block 切片」，只能让页面也交计划——`kind: 'bytes'` 就是它的计划形态。

## 撤销与保存

编辑器的顶层块有两个身份：

- `blockId`：编辑器身份，打开文档时生成，编辑 / 撤销 / 重做保持不变。
- `docxIndex`：保存文件里的坐标，每次保存后可能变化。

页面维护 `blockId -> docxIndex` 的 `anchorMap`。保存只更新这份映射和 `word/document.xml`，不再 `setContent` 重挂整个 PM 文档，因此：

- 保存 / 自动保存不会清空 undo / redo
- 从打开文档到关闭文档，同一套历史可以跨多次保存撤销
- Agent 写入后的保存同样保留历史
- 打开 / 新建 / 切换文档才清空历史

历史深度默认 1000 步（`UndoRedo.configure`）。

## 2. 目录

```text
docx-pro/
├─ package.json                 清单：声明位 + engine 模块表 + skill
├─ vite.config.ts               页面构建（root = src/page，GenOffice 走 vendor 别名）
├─ scripts/
│  ├─ build.mjs                 esbuild（ura + engine）/ vite（page）
│  └─ verify.mjs                装载前自检（镜像宿主 parseDeclaration + toolDeclIssues）
├─ skills/docx/SKILL.md         模型读的能力说明（两条路径 + 铁律）
├─ src/
│  ├─ ura.ts                    codec / sidecar / store / presenter
│  ├─ tools/
│  │  ├─ catalog.ts                  统一命令目录（blocks / editor / actions）
│  │  ├─ commands.generated.ts       10 个可用 action（由 AGENT_TOOLS 生成）
│  │  ├─ op-catalog.generated.ts     23 个编辑器 op（签名 + 字段）
│  │  ├─ declarations.generated.ts   32 条逐工具声明（仅 --legacy-tools 构建用）
│  │  └─ types.ts                    ToolDecl / RuntimeReport 的本地子集
│  ├─ engine/
│  │  ├─ tools.ts               四工具声明（guide / read / create / apply）
│  │  ├─ tools-legacy.ts        35 条逐工具声明（--legacy-tools 构建对照用）
│  │  ├─ plan.ts                只读工具 + docx_edit；页面工具回 page-unavailable
│  │  ├─ create.ts              Markdown → document.xml（新建）
│  │  ├─ exec.ts                唯一字节出口（bytes / new / save / noop / check / reject）
│  │  ├─ read.ts                字节侧只读实现（块/批注/脚注/修订/样式/诊断）
│  │  ├─ runtime.ts             能力报告
│  │  └─ zip.ts                 条目级 zip 重写（只换 word/document.xml）
│  ├─ shared/                   块命令引擎（解析 / 补丁 / 计划 / sync / sha256）+ 编辑器 IPC 类型
│  ├─ editor/                   编辑器副本（从 packages/docx fork；不跨包 import）
│  └─ page/
│     ├─ main.ts                会话 → 分块取字节 → 装宿主适配层 + agent 桥 → 起编辑器
│     ├─ host.ts                window.desktop 适配层（保存 = commit(plan)）
│     ├─ agent.ts               agent-request → 编辑器事务 → 编辑器 save() → commit
│     ├─ command-batch.ts       docx_apply：三组命令按序执行 + 失败整批回滚
│     ├─ block-commands.ts      块命令 → PM 事务 / executeOps
│     └─ env.d.ts               页面会话契约（get / bytes / commit / agentResult）
├─ tests/                       node --test（引擎 / 声明 / 往返 / 注册面）
└─ README.md
```

## 3. 工具面（4 个工具 / 41 条命令）

| 工具 | 只读 | 说明 |
| --- | --- | --- |
| `docx_guide` | ✅ | 命令目录：无参 → 三组 op 名；`group` → 组内 schema；`op` → 单条 schema + 示例 |
| `docx_read` | ✅ | `what`：`context` / `blocks` / `comments` / `notes` / `revisions` / `styles` / `audit` |
| `docx_create` | ❌ | 新建：`ref` 必须不存在，`content` 是 Markdown；空白内容 = 空白文档 |
| `docx_apply` | ❌ | 一批命令（≤200），按序执行；一条被拒整批回滚；`dryRun` 只校验 |

`docx_apply` 的 `op` 来自三组：

| 组 | 数量 | 页面 | 无页面时 |
| --- | --- | --- | --- |
| `blocks` | 8 | 映射成 PM 事务 | XML 切片（`{ kind: 'save' }`） |
| `editor` | 23 | `executeOps`（一批一个事务） | `command.page-unavailable` |
| `actions` | 10 | `executeTool`（HTML 片段 / 图表 / 修订） | `command.page-unavailable` |

为什么是四个工具：`docx_guide` + `docx_read` + `docx_apply` 负责已有文档，
`docx_create` 负责从零新建（Markdown → XML，再装进 `assets/seed.docx`）。
41 条命令的 schema 仍按需从目录取，工具列表不会膨胀。逐工具面（35 条声明，32.5 KB）
保留在 `--legacy-tools` 构建里做对照，命令实现完全相同。

**未接线的动作不进目录**（15 个）：批注 4、脚注尾注 4、页眉页脚、页面设置、分节符、
样式定义、水印、整篇生成——它们的执行体需要 App 会话内的 accessor，页面桥还没接。
目录里没有的 op 就是不存在，模型不会去猜。

## 4. 构建、测试、验证

```bash
pnpm --filter @workstudio/plugin-docx-pro build     # ura + engine + page（四工具面）
node scripts/build.mjs --legacy-tools               # 逐工具面（35 条声明，对照用）
pnpm --filter @workstudio/plugin-docx-pro typecheck
pnpm --filter @workstudio/plugin-docx-pro test      # 15 项（引擎 10 + 注册面 5）
pnpm --filter @workstudio/plugin-docx-pro verify    # 装载前自检
```

页面构建依赖 `vendor/genoffice`（`pnpm setup`）；编辑器源码是本包自带的副本
（`src/editor`，从 `packages/docx` fork），不跨包 import。
构建产物：

```
dist/ura.js
dist/engine/{tools,plan,exec,read,runtime}.mjs
dist/page/{index.html,assets/**}
```

## 5. 与另两个插件的关系

| 插件 | 写形态 | 页面 | 用途 |
| --- | --- | --- | --- |
| `docx`（既有） | `bytes` | 完整编辑器 | 纯人工编辑；无 agent 工具面 |
| `office-docx`（既有） | `plan` | 块列表（无分页） | 纯后台/批处理；无完整编辑器 |
| **`docx-pro`（本包）** | `plan` | 完整编辑器 | **两者合体**：人在场走编辑器，后台走块命令 |

三者声明同一个 MIME。**同一个工作区只能启用一个**：后注册者会在宿主注册表里
override 前者的 codec / presenter（迁移期请只保留 `docx-pro`）。

## 6. 已知边界（如实声明）

1. `write: 'plan'` 下页面必须交计划；`kind: 'bytes'` 是「整份字节的计划」。执行点
   对它会校验包结构、体积与 `document.xml` 存在性，但不做块级 diff。
2. 页面路径的 `maxCommitBytes` = 16 MB（宿主机型上限）；超过就 `command.too-large`。
   大文档请用 `docx_edit`（块级切片只产一个 part，体积小得多）。
3. 无页面时拿不到「当前选区」；`get_document_context` 只给块列表与统计。
4. 页面代理回给模型的只有 `applied` / `rejected`；工具的可读输出只进调用详情。
   需要模型读到内容时用只读工具（执行点侧）。
5. 编辑器工具在无页面时一律 `command.page-unavailable`，不回退 headless 写第二份
   模型（契约 §7.2）。
6. 块命令的 `setText` 会合并该段的 run（保留段属性与首个 run 的字符属性）；含域
   代码 / 书签的段落只适合改格式。
7. 表格只能删除 / 移动；单元格编辑必须走 `editor` 组的表格 op（需页面）。
8. 批注 / 脚注 / 页眉页脚 / 页面设置 / 样式定义 / 水印 / 已打开文档内的流式长写**尚未接线**：
   它们的执行体要 App 私有状态，页面桥只暴露了编辑器与序列化入口。
   新建独立文档由 `docx_create` 负责，不依赖这些 accessor。
