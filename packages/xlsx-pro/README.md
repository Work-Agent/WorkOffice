# xlsx-pro（Excel · `.xlsx`）

把两个既有实现合成**一个插件**：

- `packages/xlsx` 的**完整 Sheets 编辑器**（GenOffice renderer fork：Univer 网格、
  Ribbon、公式栏、格式、查找、筛选、透视、图表、打印）——源码复制进
  `src/editor`，插件之间不互相 import。
- `docx-pro` 的**计划型写协议**：页面交整份字节，执行点交 gateway 保存请求，
  字节只由 `exec` 产生。

- 插件 id：`xlsx-pro`
- 清单：`apiVersion: 2` / `write: 'plan'` / `content: 'pulled'` / `model: 'none'`
- 工具：**4 个**（`xlsx_guide` / `xlsx_read` / `xlsx_create` / `xlsx_apply`），
  目录里 **56 条命令**（cells 12 · structure 10 · sheets 6 · objects 28）
- 技能：`skills/xlsx/SKILL.md`

---

## 1. 一个插件、两种写来源、一个字节出口

```text
有页面（人在场）
  编辑器 Univer 事务（planFromOps / applyChangePlan）
    → 页面内 xlsx-gateway 产字节
    → plan { kind: 'bytes' }        ← 页面交整份字节
                          ↘
                          exec（校验包结构 / 体积）→ 权威层 CAS → 落盘

无页面（后台 / 批处理）
  xlsx_apply 命令批
    → 执行点 plan（命令 → gateway XlsxSaveRequest）
    → plan { kind: 'save' }
                          ↘
                          exec（planCellEditsToXlsx + assembleWithJsZip）
                            → 只重写 touchedEntries → 权威层 CAS → 落盘
```

`exec` 是契约里唯一的字节产生者；两条路径都经过它，版本、`epoch`、
撤销与冲突语义只有一份。

## 2. 目录

```text
xlsx-pro/
├─ package.json                 清单：声明位 + engine 模块表 + skill
├─ vite.config.ts               页面构建（root = src/page）
├─ scripts/
│  ├─ build.mjs                 esbuild（ura + engine）/ vite（page）
│  ├─ verify.mjs                装载前自检 + 一次真实往返
│  └─ gen-seed.mjs              生成 assets/seed.xlsx
├─ assets/seed.xlsx             新建空白工作簿模板
├─ skills/xlsx/SKILL.md         模型能力说明
├─ src/
│  ├─ ura.ts                    codec / sidecar / store / presenter
│  ├─ ura-types.ts              页面 / codec 共用类型
│  ├─ tools/catalog.ts          统一命令目录 + guide
│  ├─ engine/
│  │  ├─ abi.ts                 计划 / 执行点 ABI
│  │  ├─ tools.ts               4 个工具声明
│  │  ├─ plan.ts                只读工具 + xlsx_apply / xlsx_create
│  │  ├─ exec.ts                唯一字节出口
│  │  ├─ read.ts                字节侧只读
│  │  └─ runtime.ts             能力报告
│  ├─ shared/
│  │  ├─ workbook.ts            解析 / 投影 / 范围读取
│  │  ├─ headless.ts            命令 → XlsxSaveRequest
│  │  └─ types.ts               计划形状 / 保存请求
│  ├─ editor/                   完整 Sheets 编辑器副本
│  └─ page/
│     ├─ main.ts                会话 → 分块取字节 → 启动编辑器 → agent 桥
│     ├─ host.ts                （由 editor/workstudio-host.ts 承担）
│     ├─ agent.ts               agent-request → Univer 事务 → commit
│     └─ env.d.ts               页面会话契约
├─ tests/                       node --test（引擎 / 注册面）
└─ dist/
   ├─ ura.js
   ├─ engine/{tools,plan,exec,read,runtime}.mjs
   ├─ seed.xlsx
   └─ page/**
```

## 3. 工具面

| 工具 | 只读 | 说明 |
| --- | --- | --- |
| `xlsx_guide` | ✅ | 命令目录：无参 → 四组 op 名；`group` → 组内 schema；`op` → 单条 schema + 示例 |
| `xlsx_read` | ✅ | `context / sheets / range / formulas / find / audit` |
| `xlsx_create` | ❌ | 新建：`ref` 必须不存在；`sheets: [{ name, rows }]`，也支持 `rows` / `csv` |
| `xlsx_apply` | ❌ | 一批命令（≤200），按序执行；一条被拒整批回滚；`dryRun` 只校验 |

命令分组：

| 组 | 数量 | 无页面 | 说明 |
| --- | --- | --- | --- |
| `cells` | 12 | 大部分 ✅ | 单元格值 / 公式 / 区域 / 格式 / 查找替换 / 填充 / 复制 / 排序 |
| `structure` | 10 | ✅ | 行列增删、合并、行高列宽、隐藏 |
| `sheets` | 6 | ✅ | 增删改复制移动、隐藏 |
| `objects` | 28 | ❌ 需页面 | 图表 / 形状 / 图片 / 表格 / 透视表 / 迷你图 / 条件格式 / 数据验证 / 批注 / 筛选 / 冻结 / 保护 / 页面设置 / 定义名称 |

`convert_to_values` 需要 Univer 公式引擎，页面外回 `command.page-unavailable`。

## 4. 构建、测试、验证

```bash
pnpm --filter @workstudio/plugin-xlsx-pro build
pnpm --filter @workstudio/plugin-xlsx-pro typecheck
pnpm --filter @workstudio/plugin-xlsx-pro test
pnpm --filter @workstudio/plugin-xlsx-pro verify
```

页面构建依赖 `vendor/genoffice`（`pnpm setup`）；编辑器源码是本包自带的副本
（`src/editor`，从 `packages/xlsx` fork），不跨包 import。

构建产物：

```text
dist/ura.js
dist/engine/{tools,plan,exec,read,runtime}.mjs
dist/seed.xlsx
dist/page/{index.html,assets/**}
```

## 5. 与另两个插件的关系

| 插件 | 写形态 | 页面 | 用途 |
| --- | --- | --- | --- |
| `xlsx`（既有） | `bytes` | 完整 Sheets 编辑器 | 纯人工编辑；无 agent 工具面 |
| **`xlsx-pro`（本包）** | `plan` | 完整 Sheets 编辑器 | **合体版**：人在场走编辑器，后台走 gateway 计划 |

两者声明同一个 MIME。**同一个工作区只能启用一个**：后注册者会 override 前者的
codec / presenter（迁移期请只保留 `xlsx-pro`）。

## 6. 已知边界

1. 页面路径的 `maxCommitBytes` = 16 MB；超限时请拆文件或走无页面命令批。
2. `xlsx_read` 读落盘字节；页面有未保存修改时以页面上下文为准。
3. 无页面路径不重算公式：公式文本写入，缓存值可能保持旧值。
4. `format_range` 无页面只支持常用样式；复杂主题色/渐变/边框请打开页面。
5. 大范围操作展开超过 2,000 个单元格会被拒绝，拆成多次。
6. 结构命令与 cell 编辑同批时按命令顺序移动坐标；跨批修改请重新读取。
7. `objects` 组无页面一律 `command.page-unavailable`，不回退 headless 写第二份模型。
