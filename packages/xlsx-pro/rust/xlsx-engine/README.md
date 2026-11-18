# 上游 Rust 引擎 → wasm（`xlsx-engine.wasm`）

这个目录放**接入层的补丁**，不放引擎源码。引擎源码来自
`vendor/genoffice/apps/sheets/native/xlsx-engine`（`pnpm setup` 稀疏拉取）。

| 文件 | 作用 |
| --- | --- |
| `wasm.rs` | WASI reactor ABI：`wasi_alloc / wasi_free / call / _initialize / xlsx_protocol_version` |
| `../scripts/build-wasm.mjs` | 复制源码到 `build/xlsx-engine-wasm/` → 打补丁 → `cargo build --lib --target wasm32-wasip1` → 产物自证 → 发布 |

构建产物：`dist/engine/xlsx-engine.wasm`（**6.65 MB raw / 2.10 MB gzip**）+ `.json`（体积/sha256/导出面）。

---

## 源码版本要求（重要）

wasm 要求**纯 Rust 依赖树**。实测：

| 引擎源码 | 依赖 | wasm 能否编 |
| --- | --- | --- |
| `285540a4`（当前 vendor pin） | `ironcalc 0.7.1` → `zip 0.6.6` → `bzip2-sys`（C 源码） | ❌ 需要 clang |
| `f6c2358c`（upstream `main`） | `ironcalc 0.9.1` / `zip 9(deflate)` / `cfb 0.15` / `calamine 0.37` | ✅ 一条命令过 |

`build-wasm.mjs` 会在编译前**预检依赖**（`zip ≥ 9`、`ironcalc ≥ 0.9`、有 `cfb`）：

- 不满足 → 默认 **skip**（打出原因与出路，`pnpm build` 不会因此变红）；
- `--require` → 直接失败（CI 用这个）。

两条出路：

```bash
# A. 升级 vendor 的引擎源码到 >= f6c2358c（会牵动 fork 的 i18n：上游 packages/i18n
#    新增 tr / hu，同步 locale 后才能整包升级，见 docs/xlsx-pro-host-checklist.md §6）

# B. 不升级 vendor，只把构建指向一份满足要求的 checkout
XLSX_PRO_ENGINE_SOURCE=/path/to/genoffice/apps/sheets/native/xlsx-engine \
  node packages/xlsx-pro/scripts/build-wasm.mjs

# 或命令行等价
node packages/xlsx-pro/scripts/build-wasm.mjs --source /path/to/xlsx-engine
```

> 无论哪条路，产物 sha256 应当一致（同一版本源码 + 同一工具链）：实测 `4c4e1931…`。

---

## 构建

```bash
rustup target add wasm32-wasip1          # 或 --install-target
node scripts/build-wasm.mjs              # 增量：cargo target 缓存在 build/xlsx-engine-wasm/target
node scripts/build-wasm.mjs --json       # 机读输出（CI）
```

冷编 4 min / 增量 ~80–110 s。`build/xlsx-engine-wasm/` 已在 `.gitignore` 里。

## 产物自证（`verify.mjs` 也会查）

- 导出必须含 `memory / _initialize / call / wasi_alloc / wasi_free`；
- **不得**含 `_start`（那是 command module，`wasi.start` 只能跑一次）；
- 体积：raw ≤ 12 MB、gzip ≤ 4 MB（当前 6.65 / 2.10）；
- `xlsx-engine.json` 的 `sha256` 与实际文件一致。

## 已知落差

| 项 | 说明 |
| --- | --- |
| `_initialize` | rustc 的 wasip1 cdylib **不会**自动导出，必须由 `wasm.rs` 显式给 |
| 取消 | 单线程 reactor 没有 reader 线程，`cancel` 只能命中下一个请求 → `degraded:['cancel-unavailable']` |
| `save_archive` 入参 | `replacements/additions` 是**临时文件路径**（`EntryContent{name, contentPath}`），所以 gateway 产出的字符串要先落盘 |
