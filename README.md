# BetterGravity Doc Preview (`doc-preview`)

用于在 **Google Antigravity 2.x + BetterGravity** 工作区直接预览 Word（`.doc`、`.docx`）、PowerPoint（`.ppt`、`.pptx`）和 PDF 的本地离线插件。

## 目录结构

```text
bettergravity-doc-preview/
├── plugin.json              # BetterGravity 插件清单
├── package.json             # 工程脚本与依赖声明
├── README.md                # 工程使用说明
├── docs/
│   └── INTEGRATION.md       # 完整对接文档（宿主机制、DOM/Fiber 协议、构建与排查）
├── src/
│   └── plugin-core.js       # 插件核心源码
├── styles/
│   └── doc-preview.css      # 预览工具栏、页面容器与 PDF 文本层样式
├── scripts/
│   ├── build.mjs            # 离线打包脚本（输出到 dist/）
│   └── install.ps1          # 一键构建、安装并启用到 BetterGravity
└── dist/                    # 构建产物（可直接复制到 %APPDATA%\BetterGravity\plugins\doc-preview）
```

## 核心特性

- **100% 本地离线渲染**：内置 `JSZip`、`docx-preview`、`PDF.js (3.11.174)` 及内联 Blob Worker，不依赖任何外网 CDN。
- **默认自适应宽度**：打开文档自动匹配当前分栏可用宽度（支持拖动分栏实时自适应）。
- **鼠标定点缩放**：支持 `Ctrl + 鼠标滚轮` / 触控板捏合缩放（`25%` ~ `300%`），以指针所在位置为锚点同步调整滚动条。
- **PDF 懒加载与文字选中**：按视口可见范围分页渲染高分画布与可选中文本层。
- **非破坏式挂载**：隐藏原生占位节点而非删除 React DOM，切换标签或关闭插件时无 React 协调报错。

## 快速构建与安装

```powershell
# 1. 安装依赖（首次）
npm install

# 2. 构建到 dist/
npm run build

# 3. 校验产物语法
npm run check

# 4. 一键安装到 %APPDATA%\BetterGravity\plugins\doc-preview 并启用
npm run install:plugin
```

需要在工作区直接打开超过 20 MB 的本地文档，或启用 PowerPoint 原版预览时，可在安装时指定允许读取的目录。安装后重启 Antigravity。

```powershell
.\scripts\install.ps1 -PreviewRoots 'D:\Projects'
```

本地预览组件按路径读取已授权目录中的文档。其他目录可通过「选择文件预览」手动提供文件。

旧版 Word 预览需要指定包含 DocVortex 的 Python，可复用 MinerU 自带的 Python 环境。以下命令保留已有的预览目录设置。

```powershell
.\scripts\install.ps1 -LegacyWordPython 'C:\path\to\mineru\.venv\Scripts\python.exe'
```

重启本地预览组件或 Antigravity 后，Windows 上的已授权本地 `.doc` 优先用 Microsoft Word 在后台只读导出 PDF，按 Word 的分页显示。找不到 Word、文件带有下载标记或通过手动选取提供时，使用 DocVortex 兼容预览，排版可能有差异。预览按文件内容在内存中复用，原文件修改后会重新转换。单个 `.doc` 的转换上限为 128 MiB。

含嵌入图片的已授权本地 `.docx` 也优先使用 Word 原版预览，以保留浮动图、印章和裁切效果。其他 `.docx` 继续使用内置预览。

PowerPoint 文件在工作区内逐张显示，支持滚动、缩放和适应宽度。已授权目录里的文件优先由本机 Microsoft PowerPoint 只读导出，保留静态页面的版式、图片和图表。手动选取文件或原版导出不可用时，可通过上面配置的 Python 使用 DocVortex 兼容预览，版式可能有差异。转换上限为 128 MiB。预览仅显示静态幻灯片。

详细技术架构与宿主对接规范请参阅 [docs/INTEGRATION.md](docs/INTEGRATION.md)。
