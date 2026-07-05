# Pi for Excel

[English](./README.md) | 简体中文

> [!NOTE]
> 本简体中文指南由 AI 翻译/生成，可能存在译文问题；如与英文文档不一致，请以英文版为准。欢迎提 [Issue](https://github.com/tmustier/pi-for-excel/issues) 指正。

> 本文档是简要中文指南，仅涵盖 **Microsoft Excel 加载项**的**安装**与**模型配置**。完整功能说明、开发者文档等请参阅[英文版 README](./README.md) 与 [docs/](./docs/README.md) 目录。WPS 表格支持另见英文文档 [docs/wps-support.md](./docs/wps-support.md)。

Pi for Excel 是一款开源、多模型的 Microsoft Excel AI 侧边栏加载项,由 [Pi](https://pi.dev) 驱动。

它是一个运行在 Excel 内部的 AI 智能体:能读取你的工作簿、修改内容、进行联网研究——模型由你选择。既支持 Anthropic、OpenAI、Google Gemini、GitHub Copilot 的 API Key 或 OAuth 登录,也支持任何 **OpenAI 兼容接口**(如 DeepSeek、智谱 GLM、Ollama 本地模型等)。

**功能一览**(详见[英文版 README](./README.md#features)):

- 16 个内置电子表格工具:读写单元格、填充公式、全簿搜索、结构调整、单元格格式与条件格式、公式解释与依赖追踪、批注、自动备份等
- 多模型支持,对话中可随时切换模型
- 会话管理:每个工作簿多个会话标签页、自动保存/恢复、历史记录
- 自动上下文注入:AI 自动获知工作簿结构、当前选区和最近改动,无需手动描述
- 每次修改前自动创建检查点,出错可一键回滚
- 斜杠命令、扩展系统、联网搜索 + MCP 集成

---

## 安装

无需编程或开发工具——只需下载一个文件并添加到 Excel。

### 1)下载清单文件(manifest)

下载此文件并保存到容易找到的位置(例如桌面):

👉 **[manifest.prod.xml](https://pi-for-excel.vercel.app/manifest.prod.xml)**

<details>
<summary>备用下载链接(如上方链接无法访问)</summary>

- 最新 Release:https://github.com/tmustier/pi-for-excel/releases/latest
- 仓库直链:https://github.com/tmustier/pi-for-excel/blob/main/manifest.prod.xml

</details>

### 2)添加到 Excel

#### Windows

> Windows 桌面版 Excel 的界面会随版本/租户略有差异；以下是 Microsoft Office 加载项侧载路径。如找不到相同按钮，请参见下面的微软官方指南。

1. 打开 Excel
2. 点击 **插入(Insert)→ 我的加载项(My Add-ins)**
3. 点击 **上传我的加载项(Upload My Add-in…)**
4. 选择刚下载的 `manifest.prod.xml`
5. 在功能区点击 **Open Pi** 打开侧边栏

> ⚠️ 请务必通过 **上传我的加载项(Upload My Add-in…)** 安装。
> 不要通过 **管理 → XML 扩展包(XML Expansion Packs)** 导入——那是旧版 Excel 的功能,会对 Office 加载项清单报出误导性的证书错误。

更多细节参见[微软官方指南](https://learn.microsoft.com/zh-cn/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)。

#### macOS

1. 打开 Finder,按 **Cmd + Shift + G**(前往文件夹)
2. 粘贴以下路径并回车:
   ```
   ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   ```
3. 将 `manifest.prod.xml` 复制到该文件夹
4. 完全退出 Excel(Cmd + Q)并重新打开
5. 点击 **插入 → 我的加载项**,应能看到 **Pi for Excel**,点击一次以注册加载项
6. 在 **开始(Home)** 功能区最右侧找到 **加载项(Add-ins)** 按钮(四个橙色方块图标),点击后选择 **Pi for Excel** 打开侧边栏

> **文件夹不存在?** 先在终端(Terminal)中运行:
> ```bash
> mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
> ```
> 然后从第 3 步继续。

#### Excel 网页版(Office Online)

**开始 → 加载项 → 更多加载项 → 我的加载项 → 管理我的加载项 → 上传我的加载项**,选择 `manifest.prod.xml`。

> ⚠️ 网页版中加载项可能在若干天后消失,重新上传即可。此路径由社区贡献,未经官方完整测试。

### 3)首次运行检查

1. 打开侧边栏(**开始**功能区的**加载项**按钮 → **Pi for Excel**)
2. 连接一个模型服务(见下一节)
3. 发送测试消息,例如:
   - `我当前在哪个工作表?`
   - `总结一下我当前选中的区域`

收到回复即表示安装成功。

---

## 连接模型

### 方式一(推荐):API Key

对大多数用户来说,API Key 是最顺畅的方式,通常**无需**代理。

1. 在 Pi 中输入 `/login`(或使用欢迎页)
2. 展开某个服务商(OpenAI、Google Gemini、Anthropic 等)
3. 粘贴你的 API Key
4. 点击 **Save**

### 方式二:自定义 OpenAI 兼容网关(DeepSeek、智谱 GLM、本地模型等)

任何提供 OpenAI 兼容接口的服务都可以接入:

1. 在 Pi 中打开 `/settings`
2. 在 **Custom OpenAI-compatible gateways** 下填写:
   - **Endpoint**(接口基础地址 / base URL)
   - **Model**(模型 ID)
   - **API key**(部分本地服务可留空)
3. 保存网关后,在 `/model` 中选择该模型

常见示例(模型 ID 与地址请以各服务商官方文档为准):

| 服务商 | Endpoint | 模型 ID 示例 |
|---|---|---|
| DeepSeek | `https://api.deepseek.com` | `deepseek-chat`、`deepseek-reasoner` |
| 智谱 GLM(BigModel) | `https://open.bigmodel.cn/api/paas/v4` | `glm-4.6` 等 |
| Ollama(本地) | `http://localhost:11434/v1` | 本地已下载的模型 |

注意:

- 网关若是公网 HTTPS 地址,通常可直接连接,无需代理。
- localhost / 内网地址需经本地代理转发,启动 `pi-for-excel-proxy` 时可能需要配置目标主机策略环境变量(如 `ALLOWED_TARGET_HOSTS`、`ALLOW_LOOPBACK_TARGETS`、`ALLOW_PRIVATE_TARGETS`),详见[英文安装指南](./docs/install.md#4-connect-a-provider)。

### 方式三:OAuth 账号登录

支持 Anthropic、OpenAI ChatGPT、Google Code Assist / Antigravity、GitHub Copilot。

1. 在 `/login` 中点击 **Login with …**
2. 在弹出的浏览器窗口中完成登录
3. 返回 Excel,按提示完成剩余步骤
   - OpenAI 与 Google 的 OAuth 流程中,浏览器最后会跳到一个显示**"无法访问此网站"**的页面——这是正常现象!复制浏览器地址栏中的完整 URL,粘贴回 Pi for Excel 的提示框即可
   - 部分 Google Workspace 套餐还会要求填写 Google Cloud 项目 ID

#### OAuth 登录报 CORS / 网络错误?

Office 内嵌浏览器会拦截部分 OAuth 接口(典型报错:`Login was blocked by browser CORS`、`Load failed`、`Failed to fetch`)。解决方法是在本机运行一个本地 HTTPS 代理:

```bash
npx pi-for-excel-proxy
```

(若未安装 Node.js:`curl -fsSL https://piforexcel.com/proxy | sh`)

然后在 Pi 中打开 `/settings` → **Proxy**,启用代理并填入代理启动时打印的 HTTPS 地址(通常是 `https://localhost:3003`; 如端口被占用,会显示另一个本地端口),重试登录。详细说明与排错见[英文安装指南](./docs/install.md#oauth-logins-and-cors-proxy)。API Key 方式一般不需要代理。

---

## 常见问题(简)

- **"我的加载项"里看不到 Pi** —— 重启 Excel 再试;确认上传的是 `manifest.prod.xml`(不是 localhost 开发版清单)
- **侧边栏打开但是空白** —— 你的网络可能无法访问 `https://pi-for-excel.vercel.app`,请尝试更换网络或代理设置
- **如何更新** —— 大多数更新自动生效,关闭并重新打开侧边栏即可;极少数情况(清单变更)需重新下载并上传 `manifest.prod.xml`

更多排错项见[英文安装指南 · Troubleshooting](./docs/install.md#troubleshooting)。

---

## 更多文档(英文)

| 文档 | 说明 |
|---|---|
| [README.md](./README.md) | 完整功能介绍、开发者快速上手、架构说明 |
| [docs/install.md](./docs/install.md) | 完整安装指南 |
| [docs/integrations-external-tools.md](./docs/integrations-external-tools.md) | 联网搜索 + MCP 集成配置 |
| [docs/extensions.md](./docs/extensions.md) | 扩展开发指南 |
| [docs/security-threat-model.md](./docs/security-threat-model.md) | 安全威胁模型 |
| [docs/wps-support.md](./docs/wps-support.md) | WPS 表格支持现状与安装路径 |

## 许可证

[MIT](LICENSE) © Thomas Mustier
