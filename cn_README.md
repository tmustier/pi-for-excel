# Pi for Excel

> **🌐 [English](README.md) | [中文](cn_README.md)**

开源、多模型的 Excel 侧边栏 AI 加载项。由 [Pi](https://pi.dev) 驱动。

Pi for Excel 是一个驻留在 Excel 中的 AI 智能体。它能读取你的工作簿、执行修改、进行调研——使用你选择的任意模型。自带 API 密钥或通过 OAuth 登录 Anthropic、OpenAI、Google Gemini 或 GitHub Copilot。

---

## 功能特性

**核心电子表格工具** — 16 个内置工具，AI 可调用它们与你的工作簿交互：

| 工具 | 功能 |
|------|------|
| `get_workbook_overview` | 结构蓝图——工作表、表头、命名区域、表格、图表、数据透视表 |
| `read_range` | 以紧凑（markdown）、CSV 或详细（含格式）模式读取单元格 |
| `write_cells` | 写入值/公式，带覆盖保护和自动验证 |
| `fill_formula` | 跨区域自动填充公式（相对引用自动调整） |
| `search_workbook` | 跨所有工作表查找文本、值或公式引用 |
| `modify_structure` | 插入/删除行/列、添加/重命名/删除/隐藏工作表 |
| `format_cells` | 应用格式——字体、颜色、数字格式、边框、命名样式 |
| `conditional_format` | 添加或清除条件格式规则 |
| `trace_dependencies` | 追踪公式血统（前导引用或后续依赖） |
| `explain_formula` | 用通俗语言解释公式，附带单元格引用 |
| `view_settings` | 网格线、标题、冻结窗格、选项卡颜色、工作表可见性 |
| `comments` | 读取、添加、更新、回复、解决/重新打开单元格批注 |
| `workbook_history` | 列出/恢复工作簿修改的自动备份（保存间隔间） |
| `instructions` | AI 的持久用户级和工作簿级指导 |
| `conventions` | 可配置的格式默认值（货币、负数、零、小数位） |
| `skills` | 用于特定任务工作流的捆绑 Agent 技能 |

**多模型支持** — 使用任何支持的提供商；会话中可切换模型：
- **Anthropic**（Claude）— API 密钥或 OAuth
- **OpenAI** / **OpenAI Codex** — API 密钥
- **Google Gemini** — API 密钥
- **GitHub Copilot** — OAuth
- **自定义 OpenAI 兼容网关** — 在 `/settings` 中配置端点 + 模型 + API 密钥

**会话管理** — 每个工作簿多个会话标签页、自动保存/恢复、会话历史、使用 `/resume` 从上次中断处继续。

**自动上下文注入** — AI 在每次对话前自动接收工作簿蓝图、当前选择区域和最近的单元格变更。无需手动描述你正在查看的内容。

**工作簿恢复** — 每次修改前自动创建检查点。如果出错，可从侧边栏一键撤销。

**格式约定** — 一次性定义你的风格（货币符号、负号样式、小数位数），AI 自动遵循。

**斜杠命令** — `/model`、`/login`、`/settings`、`/rules`、`/extensions`、`/tools`、`/export`、`/compact`、`/new`、`/resume`、`/history`、`/shortcuts` 等。

**扩展** — 通过聊天安装侧边栏扩展（迷你应用）。AI 可直接通过 `extensions_manager` 工具生成并安装扩展代码。扩展默认在 iframe 沙盒中运行。

**集成** — 可选的外部工具集成：
- **网络搜索**（默认 Jina，可选 Serper/Tavily/Brave）+ `fetch_page` — 无需离开 Excel 即可查找和阅读外部来源
- **MCP 网关** — 连接到用户配置的 MCP 服务器以访问自定义工具

**桥接 + 高级控制**（通过 `/experimental` 管理）：
- Tmux 桥接设置——配置桥接 URL/令牌并运行健康检查
- Python / LibreOffice 桥接设置——配置桥接 URL/令牌
- 文件工作区写入/删除门控——跨会话共享工件存储（内置助手文档在 `assistant-docs/` 下始终可读）
- 高级扩展控制——远程 URL 选择加入、权限强制执行、沙盒回滚和 Widget API v2

（网络搜索和 MCP 在 `/tools` 或 `/extensions` → 连接中管理。）

---

## 🌍 语言支持

本项目在英文基础上增加了对**简体中文（zh-CN）** 的支持：

- **中文界面翻译**：大部分界面文本已翻译为中文。可在 **设置 → 高级 → 语言** 中切换，或在欢迎页使用语言选择按钮。
- **翻译系统**：所有面向用户的 UI 字符串使用轻量级 `t()` 翻译函数，零第三方依赖。翻译文件存储在 `src/language/locales/` 目录下，为 JSON 格式。
- **添加新语言**：在 `src/language/locales/` 中创建 `{lang}.json` 文件，并在 `src/language/index.ts` 中导入即可。

---

## 安装

1. 下载 [`manifest.prod.xml`](https://pi-for-excel.vercel.app/manifest.prod.xml)
2. 将其添加至 Excel——详见 [**安装指南**](docs/install.md)（含 macOS + Windows 逐步说明）
3. 单击功能区中的 **打开 Pi**
4. 连接提供商（API 密钥或 OAuth），或在 `/settings` 中配置自定义 OpenAI 兼容网关
5. 开始对话——试试"我有哪些工作表？"或"总结我当前选中的内容"

---

## 开发者快速入门

### 前置条件

- **Node.js ≥ 20**
- **mkcert** — 用于本地 HTTPS（Office.js 要求）

### 设置

```bash
git clone https://github.com/tmustier/pi-for-excel.git
cd pi-for-excel
npm install

# 生成本地 HTTPS 证书（Office.js 需要 HTTPS）
mkcert -install   # 一次性 CA 设置
mkcert localhost   # 创建 localhost.pem + localhost-key.pem
mv localhost-key.pem key.pem
mv localhost.pem cert.pem
```

### 运行

```bash
npm run dev        # Vite 开发服务器，https://localhost:3000
```

然后将开发清单侧载到 Excel：

**macOS**（[Microsoft 文档](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)）：
```bash
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```
然后打开 Excel → **插入** → **我的加载项** → **Pi for Excel**。

**Windows**（[Microsoft 文档](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)）：

打开 Excel → **插入** → **我的加载项** → **上传我的加载项** → 选择 `manifest.xml`。

开发清单指向 `https://localhost:3000`。生产清单（`manifest.prod.xml`）指向托管的 Vercel 部署。

### 有用命令

| 命令 | 说明 |
|------|------|
| `npm run dev` | 启动 Vite 开发服务器（端口 3000，HTTPS） |
| `npm run build` | 生产构建 → `dist/` |
| `npm run check` | 代码检查 + 类型检查 + CSS 主题检查 |
| `npm run typecheck` | 仅 TypeScript 类型检查 |
| `npm run lint` | ESLint |
| `npm run test:models` | 单元测试——模型排序 |
| `npm run test:context` | 单元测试——工具、上下文、会话、扩展、集成 |
| `npm run test:security` | 安全策略测试——代理、CORS、沙盒、OAuth |
| `npm run proxy:https` | OAuth 流程的 CORS 代理（默认 `https://localhost:3003`） |
| `npm run validate` | 验证 Office 加载项清单 |

### CORS 代理

某些 OAuth 令牌端点在 Office WebView 中被 CORS 阻止。如果 OAuth 登录失败：

1. 用户设置命令：`npx pi-for-excel-proxy`（如果缺少 Node.js，可用 `curl -fsSL https://piforexcel.com/proxy | sh`）
2. 开发/源码设置命令：`npm run proxy:https`（默认 `https://localhost:3003`）
3. 在 Pi → `/settings` → **代理** → 启用并设置 URL
4. 重试登录

API 密钥认证通常无需代理即可使用。

### 本地桥接（Python / tmux）

使用一键本地桥接助手：

- Python / LibreOffice 桥接：`npx pi-for-excel-python-bridge`（默认 URL `https://localhost:3340`，真实模式）
- tmux 桥接：`npx pi-for-excel-tmux-bridge`（默认 URL `https://localhost:3341`，真实模式）

在 Pi 中，这些 localhost 桥接 URL 默认使用。仅当需要非默认 URL 时才配置 `/experimental ...-bridge-url`。

真实模式前置条件：

- `python3` 必须安装（用于 `python_run` / `python_transform_range`）
- LibreOffice（`soffice` 或 `libreoffice`）是 `libreoffice_convert` 所必需的
- `tmux` 是 tmux 桥接真实模式所必需的

可选的辅助安装（macOS/Homebrew）：

- `npx pi-for-excel-python-bridge --install-missing`
- `npx pi-for-excel-tmux-bridge --install-missing`

手动 macOS 安装：

```bash
brew install tmux
brew install --cask libreoffice
```

要强制使用安全模拟模式：

- `PYTHON_BRIDGE_MODE=stub npx pi-for-excel-python-bridge`
- `TMUX_BRIDGE_MODE=stub npx pi-for-excel-tmux-bridge`

源码检出替代方案仍可通过 `npm run python:bridge:https` 和 `npm run tmux:bridge:https` 使用。

---

## 架构

Pi for Excel 是一个单页 Office 任务窗格加载项，使用以下技术构建：

- **[Vite](https://vite.dev/)** — 开发服务器 + 生产打包器
- **[Lit](https://lit.dev/)** — 侧边栏 UI 的 Web 组件
- **[pi-agent-core](https://www.npmjs.com/package/@earendil-works/pi-agent-core)** — 代理运行时（工具循环、流式传输、状态管理）
- **[pi-ai](https://www.npmjs.com/package/@earendil-works/pi-ai)** — 多提供商 LLM 抽象（Anthropic、OpenAI、Google、GitHub Copilot）
- **[pi-web-ui](https://www.npmjs.com/package/@earendil-works/pi-web-ui)** — 共享 Web UI 组件（消息渲染、存储、设置对话框）
- **[Office.js](https://learn.microsoft.com/en-us/office/dev/add-ins/)** — Excel 工作簿 API

### 源码布局

```
src/
├── taskpane/          # 应用初始化、会话管理、标签布局、上下文注入
├── taskpane.html      # 入口 HTML（加载 Office.js + taskpane.ts）
├── taskpane.ts        # 入口脚本
├── boot.ts            # 挂载前设置（CSS、补丁）
├── tools/             # 16 个核心工具 + 功能标记工具 + 注册表
├── prompt/            # 系统提示构建器
├── context/           # 工作簿蓝图缓存、选择/变更追踪
├── auth/              # OAuth 提供商、API 代理、凭据恢复
├── models/            # 模型排序 + 版本评分
├── ui/                # 侧边栏组件、工具渲染器、主题 CSS
│   └── theme/         # 设计令牌、组件样式（DM Sans + 青绿色调色板）
├── commands/          # 斜杠命令注册表 + 内置命令
├── extensions/        # 扩展存储、沙盒运行时、权限
├── integrations/      # 网络搜索 + MCP 网关集成目录
├── skills/            # Agent 技能目录 + 运行时加载器
├── experiments/       # 功能标记定义 + 切换逻辑
├── workbook/          # 工作簿标识（哈希）、会话关联、协调器
├── conventions/       # 格式默认值（货币、负数、小数位）
├── rules/             # 持久化用户/工作簿规则存储
├── compaction/        # 自动压缩阈值 + 逻辑
├── storage/           # IndexedDB 初始化
├── files/             # 文件工作区（始终可读/列表；写入/删除受功能门控）
├── audit/             # 工作簿变更审计日志
├── messages/          # 消息转换助手
├── debug/             # 调试模式工具
├── stubs/             # CSP/仅 Node 依赖的浏览器存根（Ajv、Bedrock、stream 等）
├── compat/            # 兼容性补丁（Lit、marked、模型选择器）
├── language/          # 翻译系统（t() 函数 + en.json + zh-CN.json）
└── utils/             # 共享助手（HTML 转义、类型守卫、错误处理）

scripts/               # 开发助手——CORS 代理、tmux/Python 桥接、清单生成
pkg/proxy/             # 可发布的 npm CLI 包：`pi-for-excel-proxy`
pkg/python-bridge/     # 可发布的 npm CLI 包：`pi-for-excel-python-bridge`
pkg/tmux-bridge/       # 可发布的 npm CLI 包：`pi-for-excel-tmux-bridge`
tests/                 # 单元测试 + 安全测试（约 50 个测试文件）
docs/                  # 当前文档（安装/部署/功能/策略）+ archive/ 历史规划
skills/                # 捆绑的 Agent 技能定义（web-search、mcp-gateway、tmux-bridge、python-bridge）
public/assets/         # 加载项图标（16/32/80/128px）
```

### 关键设计模式

- **工具注册表为单一真实来源** — `src/tools/registry.ts` 定义所有核心工具名称和构造。UI 渲染器、输入人性化器和提示文档均由此派生。
- **工作簿协调器** — 按工作簿序列化可变工具调用，防止多个会话标签页并发写入。
- **自动上下文** — 在每次用户消息前注入工作簿蓝图、选择状态和最近变更，使 AI 始终知道它在查看什么。
- **执行策略** — 每个工具被分类为 `read/none` 或 `mutate/content|structure`，以确定锁定和检查点行为。
- **恢复检查点** — 修改操作在写入前自动快照受影响的单元格，实现一键回滚。
- **扩展沙盒** — 不受信任的扩展（内联代码、远程 URL）默认在 iframe 沙盒中运行；内置/本地模块在主机上运行。

---

## 部署

生产构建是部署到 [Vercel](https://vercel.com) 的静态站点。维护者设置请参阅 [docs/deploy-vercel.md](docs/deploy-vercel.md)。

用户通过下载 `manifest.prod.xml` 并在 Excel 中上传来安装——清单指向托管的 Vercel URL。更新是自动的（关闭并重新打开任务窗格）。

---

## 文档

| 文档 | 说明 |
|------|------|
| [docs/install.md](docs/install.md) | 非技术用户安装指南 |
| [docs/deploy-vercel.md](docs/deploy-vercel.md) | 托管部署（Vercel） |
| [docs/extensions.md](docs/extensions.md) | 扩展开发指南 |
| [docs/integrations-external-tools.md](docs/integrations-external-tools.md) | 网络搜索 + MCP 集成设置 |
| [docs/security-threat-model.md](docs/security-threat-model.md) | 安全威胁模型 |
| [docs/compaction.md](docs/compaction.md) | 会话压缩（`/compact`） |
| [src/tools/DECISIONS.md](src/tools/DECISIONS.md) | 工具行为决策日志 |
| [src/ui/README.md](src/ui/README.md) | UI 架构 + Tailwind v4 说明 |

---

## 致谢

- [Pi](https://github.com/badlogic/pi-mono) by [@badlogic](https://github.com/badlogic)（Mario Zechner）— 驱动此项目的代理框架。Pi for Excel 使用 pi-agent-core、pi-ai 和 pi-web-ui 实现代理循环、LLM 抽象和会话存储。
- [whimsical.ts](https://github.com/mitsuhiko/agent-stuff/blob/main/pi-extensions/whimsical.ts) by [@mitsuhiko](https://github.com/mitsuhiko)（Armin Ronacher）— 轮播的"工作中…"消息改编自他的 Pi 扩展，为电子表格/财务场景重写。

---

## 许可证

[MIT](LICENSE) © Thomas Mustier
