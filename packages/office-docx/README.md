# office-docx（Word · `.docx`）

按 `genoffice/docs/genoffice-plugin.md` v1.14 与 `genoffice/docs/genoffice-plugin-docx.md` v1.1
**重新创建**的 docx 插件：计划型写（`write: 'plan'`）· 页面自取内容（`content: 'pulled'`）·
执行点无常驻模型（`model: 'none'`）· 提交后 `sync.manifest` rebase（I15）。

> 与仓库里既有的 `packages/docx`（v1 整份字节写、把 GenOffice 编辑器整包 fork 进页面层）
> **并存、互不影响**。同时装载两个插件时两者都声明同一个 MIME，注册表会记
> `codec.override` / `presenter.override`（后注册者胜出）——迁移期请只启用一个。

- 插件 id：`office-docx`
- 工具：`docx_guide` / `docx_apply`
- 宿主契约：`apiVersion: 2`（`docs/host-plugin-contract.md`）

## 1. 目录

```text
office-docx/
├─ package.json            清单：声明位 + engine 模块表（page → dist/page-v2/index.html）
├─ scripts/build.mjs       esbuild 产物（ura / engine / page / lib）；page-v2 由 vite 构建
├─ scripts/check-slicing.mjs  docx-engine ↔ parseDocxBytes 的块切分一致性探针
├─ scripts/install.mjs     复制 package.json + dist 到 WorkStudio 插件目录
├─ src/
│  ├─ shared/              页面与执行点**同一份**实现（构建期同源，I17）
│  │  ├─ types.ts          SavePlan（含 parts）/ 命令词汇 / manifest 形状
│  │  ├─ xml.ts           最小 XML 切片扫描（不建 DOM）
│  │  ├─ xml-ops.ts        段落 / run 属性补丁（含区间格式与区间改字）
│  │  ├─ document.ts       zip → word/document.xml → body 顶层块
│  │  ├─ ops.ts            命令批 → 读侧模型（人与 agent 同一套词汇）
│  │  ├─ plan.ts           模型 → SavePlan + 计划形状校验（sectPr 保序 / parts 预算）
│  │  ├─ plan-bytes.ts     引擎字节 ↔ 基线切片的块级 diff → SavePlan（页面保存通道）
│  │  ├─ zip-parts.ts      zip 部件指纹 / 差异 / 按需读取（P5 整包通道）
│  │  ├─ agent-commands.ts 宿主批次解包（工具包装 / 命令数组 → DocxCommand[]）
│  │  ├─ blocks-edit.ts    跨块编辑（旧块列表页用；agent / 渲染器路径不需要）
│  │  ├─ sync.ts           manifest 产出 + 页面 applyManifest
│  │  ├─ project.ts        sidecar 投影 / store 查询
│  │  ├─ catalog.ts        命令目录 + 三级 guide
│  │  ├─ sha256.ts         同步 SHA-256（页面 / 执行点同结果）
│  │  └─ runtime.ts        RuntimeReport 取值 + 提交上限常量
│  ├─ engine/
│  │  ├─ plan.ts           engine.plan：命令 → SavePlan（guide 也走这里）
│  │  ├─ exec.ts           engine.exec：SavePlan → 字节（唯一字节产生者）
│  │  ├─ read.ts           engine.read：project / query
│  │  ├─ runtime.ts        探活 → RuntimeReport
│  │  └─ zip.ts            条目级 zip 重写（正文 + parts；未改条目逐字节拷贝）
│  ├─ main/tools.ts        ToolDecl[]（主进程只声明，无执行函数）
│  ├─ ura.ts               codec / sidecar / store / presenter
│  ├─ page-v2/             ★ 页面（清单指向它）：渲染器 + plan 写盘 + agent 代理
│  │  ├─ main.tsx          get/bytes → blobUrl → 启动渲染器 + bus 桥（保存握手 / 外部变更）
│  │  ├─ host.ts           window.desktop 适配层：plan 提交 / 回执 / 重载
│  │  ├─ agent.ts          agent 代理：context 投影 + 命令 → PM 事务 / ops 引擎
│  │  ├─ bytes.ts          页面会话的分块字节读取（1 MB）
│  │  └─ genoffice/        复制的 docs 渲染器：editor/（整份 fork）+ shared/（ipc 类型）
│  └─ page/                旧块列表页（保留作回退；清单已不再引用）
├─ skills/office-docx/SKILL.md
└─ tests/                  node --test（解析 / 命令 / 计划 / 字节 diff / sync / zip / 执行点 / URA）
```

## 2. 构建、测试、安装

本包已在根 `pnpm-workspace.yaml` 的 `packages/*` 里，根 `pnpm install` 后即可构建
（`fflate` / `esbuild` / `typescript` 走 workspace 依赖）。

```bash
cd packages/office-docx
pnpm build             # node scripts/build.mjs → dist/{ura.js,engine/*.mjs,page/**,lib/**}，随后 vite → dist/page-v2/（渲染器版页面）
pnpm build:page-v2     # 只构建渲染器版页面（P1 只读；切换前 dist/page 仍是块列表版）
pnpm check:slicing     # 切分一致性校验：docx-engine blocks ↔ parseDocxBytes（见 docs/page-renderer-migration.md）
pnpm test              # 87 项（构建 + node --test）：见 §6
pnpm typecheck         # tsc --noEmit
pnpm verify            # 装载前自检：清单 / 声明位 / 产物 / 工具声明 / 能力报告
pnpm install:plugin    # → <userData>/plugins/installed/office-docx
```

> `install:plugin` 不叫 `install`：`install` 是 npm 的生命周期钩子名，
> 会被包管理器在安装时自动执行。

开发期也可以不复制：WorkStudio「设置 → 插件 → 安装来源」填本目录（原地加载），
宿主按 `dist/` 里的产物装载，改完源码重新 `node scripts/build.mjs` 并在插件列表
点「重新加载」即可。

## 3. 与设计文档的对应

| 设计条款 | 落点 |
| --- | --- |
| §3.1 I5 / §4.0 外壳统一、骨架不统一 | `SavePlan` 是 docx 的格式内部计划；宿主只透传 |
| §2 页面层 | 复制的 GenOffice Docs 渲染器（Ribbon + 分页 + 表格/图片/批注），`src/page-v2/genoffice/` |
| §9.5 plan 来源 | 渲染器跑自己的 `saveDocx` 拿字节 → 页面按**块级 diff** 还原 `SavePlan`（`shared/plan-bytes.ts`） |
| 正文以外的写盘 | `SavePlan.parts`（P5 整包通道）：media / rels / 页眉页脚 / 脚注 / 批注 / 编号 / 主题 |
| §4.2.1 令牌：`commit:version` / `inPage:docRev+epoch` / `baselineSync:manifest` | `src/shared/runtime.ts`；页面在 `page-v2/host.ts` 维护 `version` / `epoch` / `docRev` |
| §4.6 执行点：只写 `ctx.tempDir`、无目标写盘 | `engine/exec.ts` 只回 `bytes`，从不碰目标路径 |
| §8 / I12 写盘 CAS | 每次 `commit` 带 `baseVersion`；`stale` 后页面提示"文件已被外部修改"，不覆盖别人的文件 |
| I13 新鲜度按坐标系 | agent 批次校验 `epoch`（不符 → `command.stale-index`） |
| I14 计划不携带跨边界坐标 | 页面在本地解析命令后提交 `plan`；agent 的命令在页面会话内解析 |
| I15 提交后 baseline sync | `exec` 顺带产出 `manifest`；页面用「本次计划产出字节」当新基线（块切片与落盘一致），回执更新 `epoch` |
| I16 撤销口径唯一 | Ctrl+Z = 渲染器的 PM 历史；`exec` 对 `command:'undo'` 回 `command.unsupported` |
| §9.5 提交形状 | `payload.type = 'plan'`，`plan = SavePlan`；`applied` 取命令条数 |
| §10.2 自动保存节奏 | 渲染器自身的自动保存（plan 通道）；关闭守卫 `action:save` → `page-v2/main.tsx` 的 `save-result` 握手 |
| §10.4 失败语义 | `stale` 停自动保存并提示；`syncing` 排队重试；3 轮不收敛提示 |
| §10.5 页面代理 | `agent-request(kind: context/apply)` → `page-v2/agent.ts`：文本/结构走 PM 事务、格式走渲染器 ops 引擎 → 保存（plan）→ 回执 |
| §11.2 逆计划 | `inversePlan()` 保留（历史版本面板用）；页面 Ctrl+Z 用 PM 历史 |
| §11.6 manifest | `src/shared/sync.ts`：`{docxIndex, hash, xml?}`，对不上即全量重解析 |
| §14.1 两侧同源 | 同一份 `shared/*` 打进 `engine/*.mjs` 与 `page/assets/index.js` |

## 4. 命令词汇（catalog = 可用面）

`docx_guide` 三级发现：`{ref}` → 分组；`{ref, group}` → schema；`{ref, op}` → schema + example。
未实现的 op 不进目录（也不在本文里承诺）。

| 分组 | op | 说明 |
| --- | --- | --- |
| blocks | `insertParagraph` / `deleteBlock` / `moveBlock` | 顶层块结构（`index` = `docxIndex`） |
| text | `setText` / `findReplace` | 段落文本：带 `start`/`end` 时只替换段内区间（区间外 run 逐字节不变）/ 全文选择器替换 |
| format | `setHeadingLevel` / `setTextStyle` | 标题级别、粗斜体下划线、字号、颜色、对齐；`setTextStyle` 带 `start`/`end` 时只改段内区间（run 级格式） |
| check | `check` | 只读诊断（`audit`），不落盘 |

`rejected[]` 精确到命令下标与错误码：`block.not-found` / `block.hidden` /
`block.not-paragraph` / `command.not-found` / `command.invalid` / `command.too-large`。
批级原子（`isolation: 'batch'`）：一条被拒 → 整批不落盘。

## 5. 字节保真

- `SavePlan.saveBlocks` 是**最终顶层块序列**；未改块是 `{kind:'original', docxIndex}`，
  由 `exec` 用基线切片原样回写（不重新序列化）。
- `engine/zip.ts` 直接操作 zip 结构：**计划涉及的条目之外**，local record 与
  central directory record **逐字节拷贝**（只改 CD 里的 `localOffset`）；被替换 / 新增的
  条目重新 raw-deflate + 重算 CRC。测试用 `zipEntryRecords()` 断言未改条目的原始压缩流完全相等。
- **P5 整包通道**：`SavePlan.parts` 可替换 / 新增 / 删除正文以外的 part（图片与 rels、
  页眉页脚、脚注尾注、批注、编号、样式表、主题、文档属性）。XML 走 `xml`、二进制走
  `base64`；预算按 16 MB 提交上限折算；路径做穿越校验，`word/document.xml` 必须走块序列。
- 已知规范化：`w:body` 顶层块之间的排版空白（以及 body 级注释 / CDATA / 处理指令）在
  保存时被归一化 / 丢弃；块切片本身逐字节不变。zip64 与加密文档直接拒绝，不产出半成品。
- 工具入参 `dryRun` 由 plan 带进计划（`plan.dryRun`）后再由 exec 识别：宿主当前
  不把工具入参的 `dryRun` 透传给 exec，所以不能依赖宿主的 dryRun 通道。

## 6. 测试覆盖（`node --test`，87 项）

| 文件 | 断言 |
| --- | --- |
| `parse.test.mjs` | 顶层块 / 文本 / 标题 / 尾部 sectPr 隐藏 / 坏包诊断 / 切片回拼只差空白 / hash 稳定 |
| `ops.test.mjs` | 文本改写保属性、结构命令、格式命令、选择器命令、批级原子 + 错误码、sectPr 兜底 |
| `agent-commands.test.mjs` | 宿主批次解包：直接命令 / 工具包装 / 混合 / 垃圾输入 |
| `plan-bytes.test.mjs` | 字节 diff → 块级计划（未改 / 改写 / 删除 / 插入 / 混合）、body 外层变化拒绝、非法字节拒绝、计划 → exec → 字节等价、zip 部件闸门 + 功能名提示、布局（sectPr）与表格组落盘、sectPr 保序硬校验 |
| `blocks-edit.test.mjs` | 段首/段尾合并、跨段选区替换（含表格拒绝）、多行粘贴分段与光标端口、产出的仍是普通 `SavePlan` |
| `runs.test.mjs` | run 字符区间与 `paragraphTextOf` 同口径、区间格式只改覆盖的 run、区间改字沿用前一个 run 的 rPr、空段落退化 / 非法区间拒绝 |
| `sync.test.mjs` | 计划 → document.xml → manifest → 页面 rebase 后再生成全 original；对不上即 mismatched |
| `zip.test.mjs` | CRC 向量；只换一个 part，其余压缩流逐字节不变；缺 part 报错 |
| `engine.test.mjs` | guide 三级、命令 → 计划、rejected、exec 字节 / dryRun / check / undo 拒绝、project / query、runtime、**页面与执行点同一批 → 同一计划** |
| `roundtrip.test.mjs` | **两次连续提交**：manifest rebase 后用新坐标写对块；未改块切片逐字节不变；check 不产字节 |
| `ura.test.mjs` | codec 解码 / `encode` 禁止 / sidecar 索引即命令目标 / store / presenter 不放内容 |

## 7. 已知差异与未实现（如实声明，不做静默降级）

1. **页面层是 GenOffice Docs 渲染器的「副本」**：`src/page-v2/genoffice/editor` 是从
   `packages/docx/src/editor` 整份复制进来的 fork（选择复制而非抽公共包，见
   `docs/page-renderer-migration.md`）。因此：
   - 渲染器补丁**只在本包内打**，不回灌 v1；反过来，v1 那边修了 CSS/交互也要手动同步过来
     （已知一次：抑制 WorkStudio `appearance.css` 焦点环的 15 行 CSS）；
   - 分页 / 行盒 / 字体度量 / 表格 / 图片 / 批注都由渲染器提供，不再是本包的工作。
2. **没有 ProseMirror 模型**：命令直接作用于块模型，但**字符格式已到 run 级**：
   `setTextStyle` / `setText` 的 `start` / `end` 是段内字符偏移（与
   `paragraphTextOf` 同口径），页面把 DOM 选区映射成区间命令，于是“选中半段加粗”
   与“打字不冲掉区间格式”都成立（`xml-ops.ts` 的 `runSpans` / `splitRunAt` /
   `applyRunStyle` / `replaceRunText`）。仍不支持：`w:rStyle` 字符样式继承、
   修订前像（`w:rPrChange`）、跨 run 的域代码（`w:instrText` 不计字符宽度）；
   不带区间的 `setText`（agent 路径）仍会合并整段 run。
   **跨块编辑**：渲染器路径（人编辑）由 ProseMirror 原生处理；`shared/blocks-edit.ts`
   只服务旧块列表页与无页面 agent 路径（段首退格合并 / 跨段选区替换 / 多行粘贴分段）。
3. **未实现的 op 不出现在 catalog**：图片 / 图表 / 文本框 / 批注 / 修订 / 页眉页脚 /
   分节 / TOC / 样式 upsert / 邮件合并。表格只支持删除与移动，单元格编辑留给
   P1（需要 `tbl` 内部坐标）。
4. **标题样式不做 upsert**：`setHeadingLevel` 写 `pStyle HeadingN` + `outlineLvl`，
   依赖文档已有对应样式定义；没有定义时 Word 会按正文显示。
5. **无页面路径不提供 undo**（按设计 I16）；`inversePlan()` 已备好，供二期历史面板。
6. **zip64 / 加密文档拒绝**（`command.unsupported` / `command.invalid`），不写半成品。
7. **恢复副本未实现**（宿主侧能力：`docs/host-plugin-contract.md` §2 标 ❌）；页面只在
   关闭守卫里 flush，崩溃兜底依赖宿主后续落地。
8. **计划不是从 PM 文档直接映射来的**：渲染器的 `pmDocToSavePlan` 会把改过的块标成
   `generated`（结构化块），序列化需要引擎内部的 `GenerateContext`（样式表 / 编号 /
   rels 分配器），页面拿不到。所以页面**照旧跑渲染器的 `saveDocx` 拿到字节**，再按
   「新 document.xml 切片 ↔ 基线切片」做块级 diff 还原 `SavePlan`（`shared/plan-bytes.ts`），
   正文以外的变化走 `parts`（P5）。
   代价：每次保存多一趟序列化（`saveDocx` 的 CPU 开销，v1 本来就有）；
   收益：`exec` 零改动、字节等价性可单测钉住。
   **性能**：页面主 chunk 约 4.5 MB（渲染器整包，含 TipTap / 分页 / i18n / 字体），
   尚未做路由级代码分割 —— 见 §8 的 P6 待办。
9. **外部变更会清空本地历史**：重新载入后撤销栈从头开始（设计允许「外部变更断点」，
   本版选择更简单的「重载即重置」）。
10. **主进程解析**：宿主 SDK 目前没有 `engine.call`，`ura.decode` / `sidecar` 只能在主
   进程内联解析（只解 `word/document.xml`，投影截断到 4000 块 / 1 MB 文本）。
   `RuntimeReport.execution = 'isolated'` 描述的是 plan / exec 的落点。
11. **写盘能力边界 = 提交上限**：正文以外的 part 走 P5 通道后，凡渲染器能序列化的都能落盘；
   剩下的限制是**单次提交 16 MB**（`MAX_COMMIT_BYTES`，既约束产出也约束 parts 预算）、
   zip64 / 加密文档拒绝。大图片 / 大量附件会在这里被拦下并给出明确提示。
12. **agent 批次的原子性是「先校验再执行」**：整批命令先做形状与坐标校验（任一非法则整批不执行），
   但不做回滚式事务；`findReplace` 走渲染器的选择器语义（忽略 `all`，命中即全替换）；
   同一批里混合「文本改写」与「格式」时，格式作用于改写后的文本。
13. **外部变更的处理口径**：页面没有本地改动时整篇重载（走渲染器的打开通道）；
   有本地改动时不动，交给下一轮 `commit` 的 `baseVersion` CAS 报 `command.stale`。

## 8. 里程碑状态（对照 `genoffice-plugin-docx.md` §5）

| 里程碑 | 状态 | 说明 |
| --- | --- | --- |
| M0 宿主契约 | 依赖宿主 | 计划型写 / CAS / 页面通道 / 执行点 / 工具面 / 页面代理已接线 |
| M0.5 RuntimeReport | ✅ | `engine.runtime` 探活 + 三处同源 |
| M0.7 缓存预算 | ✅（不适用） | docx 执行点无常驻模型；页面模型随页面生死 |
| M0.9 令牌 + baseline sync | ✅ | `epoch` 镜像、`stale-index`、`manifest` / `applyManifest` |
| M1 预览 | ✅ | 工作区点开即渲染器版预览：分页 / 行盒 / 字体度量 / 页眉页脚 / 表格 / 图片 |
| M2 命令 / 计划 / 逆计划 | ✅ | 无页面执行点 headless 跑同一套命令；逆计划已备好 |
| M3a 自动保存 | ✅ | 400 ms / 2 s / 失焦 / Ctrl+S / 关闭守卫 / 外部变更冲突 |
| M3b-1 页面历史撤销 | ✅ | 渲染器的 PM 历史（Ctrl+Z / 重做），提交不清空 |
| M3b-2 逆计划 undo | ⏸ | 页面用 PM 历史；`exec` 对 undo 明确 unsupported |
| M5 技能 | ✅ | `skills/office-docx/SKILL.md`（≤60 行） |
| 页面层迁移（渲染器） | ✅ P0–P5 | 抽/复制渲染器 → 只读挂载 → plan 写盘 → agent 代理 → 布局/表格 → **P5 整包 part 通道**（图片/页眉/批注/编号）|
| P6 收尾 | 部分 | 文档与清理已完成；**路由级代码分割**（4.5 MB 主 chunk）与真机回归待做 |
