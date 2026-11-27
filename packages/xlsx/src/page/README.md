# page

`?ref=` → `page status/open` → 安装 `window.__xlsxPlus`（编辑器唯一的 I/O 出口）→
取字节、本地解析出工作表快照 → 注入 `__xlsxProBoot` → 启动 `@sheets/main`。

- 取字节：`page {op:'chunk'}` 分块（1 MiB/块，按 `hash` 绑定）；
- 保存：`page {op:'save'}` 交保存请求或整份字节，令牌 + `baseHash` 把关；
- 主题：`subscribe('shell','appearance')` → `setHostTheme` → 编辑器；
- 脏态：编辑器 `onDirty` → `report`（400 ms 防抖）；
- 租约：`attach`（冲突时 `takeover:true`）+ 15 s 心跳 + `pagehide` 的 `detach`。

P2 会把「取整份字节 + 页面解析」换成 `page meta/range` 流式加载，本目录之外不用改。
