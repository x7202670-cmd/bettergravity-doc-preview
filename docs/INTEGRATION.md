# BetterGravity Office & PDF Preview 对接文档

本文档记录 `doc-preview` 插件与 **Antigravity 2.32.x（Electron 44）+ BetterGravity 3.x** 宿主的对接协议、DOM 接管机制、数据源抓取方式、离线打包方案与调试排查指南，便于后续维护、二次开发或移植。

---

## 1. 宿主运行环境与插件规范

### 1.1 插件发现与加载路径
BetterGravity 从以下全局目录加载本地插件：

- **插件目录**：`%APPDATA%\BetterGravity\plugins\doc-preview\`
- **启用配置**：`%APPDATA%\BetterGravity\settings.json` 中的 `plugins.enabled` 数组需包含 `"doc-preview"`；开启 `"developerMode": true` 后，修改插件文件可自动热重载。
- **运行日志**：`%APPDATA%\BetterGravity\runtime.log`（所有 `plugin.log.info` / `warn` / `error` 输出均会实时写入该文件）。

### 1.2 `plugin.json` 清单规范
```json
{
  "name": "Office & PDF Preview",
  "description": "在工作区预览面板中直接显示 Word、PowerPoint 和 PDF 文档。",
  "version": "1.0.0",
  "author": "BetterGravity Community",
  "main": "index.js",
  "styles": "styles/doc-preview.css"
}
```
> **注意**：BetterGravity 仅自动加载 `main` 指定的单个入口 JS 文件与 `styles` 指定的 CSS 文件，不会在渲染进程中直接解析 Node `require()` 或外部相对路径模块，因此第三方前端库需在构建期打包进 `index.js`。

### 1.3 使用到的宿主 `plugin` API
| API | 用途 |
| :--- | :--- |
| `plugin.settings.define(schema)` | 在 BetterGravity 设置面板中注册开关与默认缩放比例（`enablePdf`、`enableDocx`、`defaultZoom`） |
| `plugin.react.getFiber(el)` | 从 DOM 节点获取关联的 React Fiber 实例（若不存在则回退读取 `__reactFiber*` 属性） |
| `plugin.log.info / warn` | 向 `runtime.log` 写入诊断与错误日志 |
| `plugin.onDispose(cb)` | 插件热重载或禁用时销毁观察器、移除挂载节点、恢复原生节点可见性、释放 Worker Blob URL |

---

## 2. 宿主 DOM 结构与接管机制

### 2.1 `.docx` 无法预览场景（已实测验证）
当在右侧 `Workspace` 点击 `.docx` 文件时，Antigravity 中间预览区会渲染以下 React 组件结构：

- 面板滚动容器：`div.flex.flex-col.h-full.w-full.pl-2.overflow-y-auto`
- 居中占位节点：`div[data-testid="binary-file-no-preview"]`
  - 文案：`Preview is not available for this file type.`
  - 按钮：`<button type="button" class="underline">Download <URL编码文件名></button>`

### 2.2 PDF 原生预览与大文件提示
当 Antigravity 提供 PDF 文件数据时，原生预览使用 `[data-testid="artifact-pdf-view"]`。插件沿该节点的 React Fiber 向上定位同时携带 `fileUri` 与 `data` 的组件，用同一组属性确定文档名称和二进制数据。PDF 文档在缩放时复用已打开的 PDF.js 实例。

文件超过 Antigravity 的 20 MB 预览上限时，插件从错误面板的 React 组件识别当前文档。安装时若通过 `-PreviewRoots` 指定了读取目录，Antigravity 会在启动时运行 `doc-preview-local` 组件，插件即可从同机取得文件数据。组件仅监听 `127.0.0.1`，校验安装时生成的随机令牌，并在解析实际路径后检查目录和扩展名。再次传入 `-PreviewRoots` 可更新目录名单。超出授权目录或读取失败时，预览面板仍提供本地选择、拖放和在文件资源管理器中显示的操作。

### 2.3 非破坏式 React DOM 挂载约定
为防止 React 在卸载或切换标签页时抛出 `removeChild` / `insertBefore` 协调异常：

- **禁止**直接 `.remove()` 或 `innerHTML = ""` 覆盖 React 管理的原有子节点；
- 对原生占位容器 `fallbackBox` 仅添加 `.bg-doc-preview-hidden-fallback`（`display: none !important`）；
- 通过 `fallbackBox.insertAdjacentElement("afterend", root)` 挂载独立的 `.bg-doc-preview-root` 兄弟节点；
- 每次 `scan()` 收集当前预览节点，检查文件标识和数据是否改变，再移除或更新对应的预览实例。

---

## 3. 文件字节流获取协议（核心对接点）

Antigravity 前端运行在 `https://127.0.0.1:<随机端口>/` 下，直接 `fetch("file:///...")` 会被 Electron 安全策略拦截（`Failed to fetch`）。实测发现宿主在渲染预览组件时，实际将文件数据缓存在上层 React 组件的 Hooks 状态中。

### 3.1 读取优先级

1. 从二进制预览占位节点或 PDF 预览组件取得当前文件的 `fileUri`。普通文件使用该组件的数据、其下载链接，或在具有相同 `fileUri` 的 React 组件上读取当前的二进制文件状态。打开旧节点时，插件会检查 React 当前提交的 Fiber。
2. 若宿主数据不可用，已授权的本地 PDF 和 DOCX 通过只读组件按 `fileUri` 读取。插件的「重载」优先读取授权目录内的当前文件。读取后检查 PDF 或 DOCX 文件头。
3. 数据引用、链接或文件状态改变时，即使宿主复用面板节点也重新加载。读取失败时提供本地选择或拖放入口。

DOCX 渲染关闭 `renderAltChunks`，文档内嵌 HTML 片段不会进入预览页面。

---

## 4. 离线打包方案说明

为避免内网/无外网环境下 CDN 加载失败，以及兼容宿主可能存在的 `module` / `exports` 全局变量，[scripts/build.mjs](../scripts/build.mjs) 采用 **UMD 隔离包装 + 内联 Worker 字符串** 的打包方式：

1. 将 `jszip.min.js`、`docx-preview.min.js`、`pdf.min.js` 分别包裹在独立的 `(function() { const module = { exports: {} }; ... return module.exports; })()` IIFE 中，避免污染全局作用域或与外部 CommonJS 环境冲突。
2. 将 `pdf.worker.min.js` 通过 `JSON.stringify` 内联为 `PDF_WORKER_SOURCE` 常量；运行时通过 `URL.createObjectURL(new Blob([PDF_WORKER_SOURCE], { type: "application/javascript" }))` 动态注入 `pdfjsLib.GlobalWorkerOptions.workerSrc`。
3. 所有代码合并输出为单文件 `dist/index.js`，不依赖 `eval` 或 `new Function`。

---

## 5. 缩放与交互实现说明

- **默认适应宽度（`fit`）**：
  - Word：在 `docx.renderAsync` 完成后测量实际页宽（默认回退 `794px`），按 `(viewport.clientWidth - 28) / naturalWidth` 计算倍率并设置 `docxStage.style.zoom`。
  - PDF：创建分页占位时记录最大页宽，再按统一倍率适应分栏。
- **鼠标定点缩放（`Ctrl + 滚轮`）**：
  - 在 `.bg-doc-preview-root` 捕获阶段监听 `wheel` 事件，拦截 `ctrlKey || metaKey` 并 `preventDefault()` + `stopPropagation()`；
  - 缩放前记录指针在舞台元素视觉矩形中的归一化相对位置 `(relX, relY)`，缩放后立即通过调整 `viewport.scrollLeft` / `scrollTop` 将该文档坐标拉回指针正下方；
  - PDF 使用 CSS `zoom` 即时缩放页面和文本层。停止缩放 `180ms` 后按 `currentZoom * devicePixelRatio` 重绘视口附近的画布，完成后才替换旧画布。画布离开视口及其预加载区域后，保留最近使用的最多 4 页且总计不超过约 64 MiB，返回已回收页面时重新绘制。

---

## 6. 二次开发与排查指南

1. **本地修改与热更新**：
   - 修改 [src/plugin-core.js](../src/plugin-core.js) 或 [styles/doc-preview.css](../styles/doc-preview.css) 后，运行：
     ```powershell
     npm run install:plugin
     ```
   - 该命令会自动重新打包 `dist/`，同步覆盖至 `%APPDATA%\BetterGravity\plugins\doc-preview\`，并在 `settings.json` 中登记 `doc-preview`。开发者模式关闭时，安装脚本会提示启用。
2. **宿主升级后排查步骤**：
   - 打开 `%APPDATA%\BetterGravity\runtime.log`，搜索 `doc-preview` 关键字；
   - 若宿主升级后普通 DOCX 无法读取，可核对 `BINARY_PREVIEW_SELECTOR` 及与同一文件绑定的 `fileContents` 或二进制 Hook 状态；
   - 若原生 PDF 组件更换了挂载容器结构，可检查 `PDF_VIEW_SELECTOR` 和 `getHostFileSource` 获取的属性。
