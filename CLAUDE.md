# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

我在windows10上面开发，使用的开发工具是Visual Studio Community 2022 + Visual Basic.NET + VSTO，使用的语言是Visual Basic.NET，使用的框架是.NET Framework 4.7.2，使用的Office插件是VSTO，使用的Office版本是Office 2016+ / WPS。

- **官网**: https://www.officeso.cn
- **License**: Apache 2.0

## Repository Structure

```
office-ai-agent/
├── ExcelAi/           # Excel VSTO 插件 (同时使用 ExcelDna 提供 UDF 函数)
├── WordAi/            # Word VSTO 插件
├── PowerPointAi/      # PowerPoint VSTO 插件
├── ShareRibbon/       # 共享核心库 - 所有 UI、服务、Agent、MCP、数据库
├── OfficeAgent/       # 安装包项目 (.vdproj)
├── AiHelper.sln       # 主解决方案文件
└── Directory.Build.targets  # 构建辅助（创建缺失的 XML 文档占位文件）
```

三个 Office 宿主项目 (ExcelAi/WordAi/PowerPointAi) 都是薄壳，通过 `ProjectReference` 引用 `ShareRibbon`。所有共享代码都在 `ShareRibbon` 中。

## Build & Development

### Prerequisites

- Visual Studio 2022 (with VSTO 工作负载)
- .NET Framework 4.7.2
- Office 2016+ / WPS

### Build Commands

```bash
# 还原 NuGet 包
nuget restore AiHelper.sln

# 构建整个解决方案
msbuild AiHelper.sln /p:Configuration=Debug

# 构建单个项目
msbuild ShareRibbon/ShareRibbon.vbproj
msbuild ExcelAi/ExcelAi.vbproj
msbuild WordAi/WordAi.vbproj
msbuild PowerPointAi/PowerPointAi.vbproj
```

### Build Gotchas

- **VSTO 清单签名**: Markdig.dll 等未签名程序集会触发 MSB3178 错误。ExcelAi 通过 `CopyUnsignedDllsAfterBuild` target 复制 Markdig.dll，通过 `ExcludeUnsignedAssembliesFromManifest` target 排除 25+ 个第三方程序集
- **vdproj 安装项目**: 极易被自动修改导致加载失败，出问题先 `git checkout` 回退
- **Directory.Build.targets**: 为缺失的 System.*.xml 文档文件创建占位，防止打包验证失败

## Key Architecture

### WebView2 UI 通信管道

这是最核心的架构模式。Chat UI 使用 WebView2 内嵌浏览器，前后端通过消息传递通信：

**VB.NET → JavaScript**: `CoreWebView2.ExecuteScriptAsync(script)` 执行任意 JS
**JavaScript → VB.NET**: `window.chrome.webview.postMessage(payload)` 触发 `WebMessageReceived` 事件

`window.vsto` 桥接对象提供结构化 API：
- `window.vsto.sendMessage(payload)` — 发送聊天消息
- `window.vsto.executeCode(code, lang, preview)` — 执行代码
- `window.vsto.saveSettings(settingsObj)` — 持久化设置

消息由 `MessageService.HandleWebMessage()` 按类型分发到对应处理器。支持的 message type: `sendMessage`, `stopMessage`, `executeCode`, `saveSettings`, `startAgent`, `abortAgent`, `agent:approvePlan`, `agent:rejectPlan`, `agent:refinePlan`, `clearContext`, `acceptAnswer`, `rejectAnswer`, `applyRevisionAll` 等。

### 前端资源加载

HTML/CSS/JS 通过 Office Virtual Server 加载。`ResourceExtractor` 在启动时将嵌入资源解压到临时目录，然后 `SetVirtualHostNameToFolderMapping` 将 `officeai.local` 映射到该目录。

**前端 JS 架构** (18 个文件，全部位于 `ShareRibbon/Resources/`):
- `utils.js`, `core.js` — 基础配置、marked.js 初始化
- `chat-manager.js`, `message-sender.js` — 聊天编排和发送
- `markdown-renderer.js` — Markdown 转 DOM 渲染
- `code-handler.js` — 代码块复制/执行/编辑
- `agent-protocol.js`, `agent-card.js` — 新 Agent 系统的前端协议和卡片 UI
- `ralph-loop.js`, `ralph-agent.js` — 旧 Ralph Loop 前端
- `mcp-manager.js` — MCP 连接管理 UI
- `history-manager.js`, `revision-manager.js` — 历史和修订
- `intent-preview.js` — 意图识别展示
- `settings-manager.js`, `model-switcher.js`, `config-panel.js` — 设置/模型/配置 UI
- `autocomplete.js`, `reformat-template.js` — 自动补全和排版模板

### 服务组合模式 (Delegate-based DI)

不使用 DI 容器。`BaseChatControl` 作为组合根，通过懒加载属性创建服务，构造函数接收委托回调（`Func(Of ...)`, `Action(Of ...)`）。几乎所有服务都遵循模式:

```vbnet
If _service Is Nothing Then _service = New XxxService(callback1, callback2)
Return _service
```

关键服务及其职责:
- **WebViewService** — WebView2 初始化、虚拟主机、脚本执行
- **MessageService** — WebMessage 消息路由
- **HttpStreamService** — SSE 流式请求、MCP tool call 拦截（ReAct 循环，最多 5 次 tool call）、Markdown 缓冲
- **CodeExecutionService** — VBA 执行、JSON 命令分发、公式求值
- **ChatContextBuilder** — 7 层上下文组装 ([0]Role [1]Scenario [1b]Skills [2]Profile [3]RAG [4]Summary [5]Query [6]ToolResults)
- **IntentRecognitionService** — 意图分类（结合 referenceSummary + ragSnippets + 当前会话）
- **McpService** — MCP 连接管理
- **AgentKernelService** — 新统一 Agent 的包装器

### 启动性能策略

`PhaseStartupManager` 实现 3 阶段启动：
- **Phase 0** (关键路径): 仅注册事件处理器
- **Phase 1** (后台): 预加载程序集
- **Phase 2** (fire-and-forget): WebView2/SQLite/资源加载

WebView2 和 SQLite 使用 `Lazy(Of Boolean)` 延迟到首次访问时才初始化。

### 双 Agent 系统

项目中有两套 Agent 系统，通过 `ConfigSettings.UseNewAgentKernel` (默认 `True`) 切换:

| 系统 | 位置 | 特点 |
|------|------|------|
| **Ralph Loop (旧)** | `ShareRibbon/Loop/` | RalphLoopController (简单任务规划) + RalphAgentController (Spec→意图→规划→执行) |
| **AgentKernel (新)** | `ShareRibbon/Agent/` | 统一 Agent，集成 6 个子系统 |

**AgentKernel 子系统**:
1. `PromptManager` — 从文件系统加载提示词
2. `ToolRegistry` — 管理可用工具 (本地 + MCP)
3. `SkillRegistry` — 管理技能
4. `AgentMemory` — 短期/长期记忆 + RAG
5. `LoopEngine` — ReAct 循环 (Think → Plan → Act → Observe → Reflect)
6. `AgentSession` — 统一数据模型

**ReAct 循环流程**: Spec 生成 → Plan 生成 → 逐步执行 → 观察反思 → 最多 15 次迭代，最多 3 次无进展循环，最多 2 次重新规划。简单任务自动执行，中等/复杂任务需用户确认。

### JSON 命令模式

新旧 Agent 系统共用 22 条 JSON 命令模式，每个 Office 应用各有一套特定实现（ExcelAi/WordAi/PowerPointAi 各有一个 `XxxDirectOperationService` + `XxxJsonCommandSchema`）。

### MCP 客户端

`StreamJsonRpcMCPClient` 支持双向传输:
- **SSE 模式**: HTTP + JSON-RPC 2.0，Bearer token 认证
- **Stdio 模式**: `stdio://` URL 自动解析为子进程通信（node/python/cmd.exe 自动检测）

连接配置存储在 `%MyDocuments%\OfficeAiAppData\mcp_connection.json`。

### 数据库

SQLite 数据库位于 `%MyDocuments%\OfficeAiAppData\office_ai.db`，通过 `System.Data.SQLite` + EF6 访问。使用 WAL 模式。

**重要**: 有 8 个迁移版本（`schema_version` 表），新增字段必须通过 `ALTER TABLE` 迁移脚本处理，不能只在新环境建表。所有 Repository 使用静态方法 + 直接 `SQLiteCommand`（无 EF 上下文）。

### 配置文件位置

所有配置文件位于 `%MyDocuments%\OfficeAiAppData\`:
- `office_ai_config.json` — API 提供商配置
- `office_ai_prompt_config_{appname}.json` — 各应用提示词模板
- `office_ai_chat_settings.json` — 聊天 UI 设置
- `mcp_connection.json` — MCP 服务器连接
- `ralph_memory.json` — Ralph Loop 记忆
- `office_ai.db` — SQLite 数据库

### ExcelAi 特有：双插件技术

ExcelAi 同时使用 **VSTO** (任务窗格、Ribbon、COM 交互) 和 **ExcelDna** (UDF/XLL 函数)。XLL 文件启动时通过多级路径搜索自动发现（注册表 → 标准安装目录 → 当前目录 → 父目录 → 全盘扫描）。

## Important Conventions & Pitfalls

1. **新增 JS/CSS/HTML**: 必须在 `ShareRibbon.vbproj` 中配置为 `None`/`EmbeddedResource`，并确保主 HTML 引用且通过 Virtual Server 可访问
2. **vdproj 安装项目**: 慎改，自动修改后易加载失败，出问题先 `git checkout` 回退
3. **数据库迁移**: 新字段必须走 `ALTER TABLE` + 版本号升级控制
4. **意图识别**: 发送前需收集 referenceSummary + ragSnippets + 当前会话上下文
5. **WPS 兼容**: 通过 `LLMUtil.IsWpsActive()` 检测，WPS 下任务窗格宽度需特殊处理
6. **选项设置**: 项目级 `Option Strict Off`，`Option Infer On`。生成的设计器文件使用 `Option Strict On`
7. **中文交互**: 请用中文与项目维护者和代码注释交互

## Reference

- **AGENTS.md**: 更详细的代码库知识库
