# OfficeKit 门禁分层

OfficeKit 把“开发反馈速度”和“发布可信度”分开。小改动先走 fast gate；会下载外部 provider、启动 LibreOffice/Poppler、运行 PromptBench、渲染完整 reference Skill 矩阵、重跑全部模板或重建 WASM 的检查，只在 slow gate、发布候选或 nightly 执行。

## 三类入口

| 层 | 本地入口 | Hosted 入口 | 触发 | 包含 |
| --- | --- | --- | --- | --- |
| Fast | `npm test`（等同 `npm run test:fast`） | `.github/workflows/ci.yml` | 每次 push/PR | JS syntax/import、核心四格式模型、OfficeKit 路由/参考插件路径同步/可移植性 validator、Help、包内容 smoke。不会下载 Python/PyMuPDF/qpdf/OCR/veraPDF，不运行完整域 Skill 或 reference Skill render/matrix、PromptBench、clean-install 或 20 套默认模板的全量原生回归。 |
| Slow | `npm run test:slow` | `.github/workflows/ci-slow.yml` | 每日定时或人工冻结 milestone；不随普通 push/PR 自动触发 | 原完整测试链、provider/pack、Playwright/LibreOffice/Poppler、PromptBench candidate/reference、默认模板全量 import/export/recalc/render、examples、release/package、OfficeBridge 与 OfficeKit WASM。Hosted workflow 将同一条步骤表按十个职责段顺序执行，避免单个长命令被 runner 取消。 |
| Windows Office live | 见下文 | `.github/workflows/windows-office-live.yml` | 手动排队；Live host 变更或 release candidate | 真实 Windows + Microsoft Office 人工观察证据。GitHub-hosted Windows、macOS mock、Add-in build smoke 和 CLI/package smoke 都不能替代这条证据。 |

`npm run test:slow:templates` 和 `npm run test:slow:promptbench` 是 slow gate 中可单独复跑的两个窄入口。它们不改变正式 slow gate 的完整范围，也不应被记录成完整发布证据。

Hosted `ci-slow` 和手动 release workflow 使用同一 runner 入口的十个连续段：

1. `foundation`：基础模型、路由、公式与 sparkline
2. `presentation`：Presentation 模型、JSX 和 Presentation Skill
3. `templates`：模板库完整性、六个默认模板 shard
4. `officekit`：Template Creator、OfficeKit Skill、CLI 和 REPL
5. `documents`：Live smoke 与 Documents 工作流
6. `pdf-packs`：PDF provider pack 构建与 managed-release
7. `pdf-providers`：基础 PDF provider 合约与编辑 provider
8. `pdf-specialists`：签名、PDF/A、OCR、Skill 与 PromptBench
9. `qa`：验证、review、render、visual baseline、renderer 和 OfficeBridge
10. `release`：examples、release/package、standalone、Help

这些段只改变 hosted 调度，不改变覆盖范围或步骤顺序；本地完整
`npm run test:slow` 仍是单一串行入口。每段都可用
`npm run test:slow -- --segment <name>` 单独复跑，未知段名会 fail closed。
其中 `templates` 内的默认模板矩阵按
`documents-a/b`、`presentations-a/b`、`spreadsheets-a/b` 六个模板 shard
顺序执行；`node test/default-template-library.mjs --shard <name>` 可单独
复跑一个 shard，未指定 shard 时仍运行完整本地矩阵。

模板库 slow 输出会标记每个模板的 materialize、roundtrip、native recalc 和 render 阶段；若 hosted runner 取消或超时，交接记录应保留最后一个模板/阶段标记，不得把整组矩阵写成“未执行”。

## 变更判断

- 每个原子提交先跑受影响的定向测试和 `npm test`；普通 Skill/JS 改动不自动升级为完整发布候选。
- proto、C# Codec 或 bundled WASM 改动另跑最窄的 `proto:check`、对应 .NET 测试和确定性 WASM build；provider、模板、PromptBench 或 Live 改动使用各自的窄入口。窄门禁结果不能写成完整发布证据。
- 同一领域累计 3–5 个已闭环纵切后冻结一个 milestone，再考虑完整发布候选。两次完整候选的**启动时间**必须至少相隔 12 小时；这是滚动窗口，不从上一轮结束时间重新计时。紧急安全修复或用户明确要求立即发布候选才可例外，并必须记录原因。
- 发布候选：冻结版本和包后一次性运行 `npm run test:slow`、`npm run docs:api`、`npm run test:pack`、`npm run release:check`、OfficeBridge/OfficeKit .NET tests，再生成 standalone/release pins，并触发一次 hosted slow。npm auth、tag、Windows 实机等外部阻塞仍单独记录。
- `docs/api.md` 只有在公开 API 改变时重生成；fast gate 的 Help/API 断言不代替发布前 `docs:api` clean diff。

## Windows Office live 证据

`.github/workflows/windows-office-live.yml` 使用 `[self-hosted, windows, office]` runner，故没有可用实机时会保持排队，而不是把 `windows-latest` 或 macOS mock 宣称为 Office 验收。操作员在真实 Excel/PowerPoint 中完成对应工作流后，提供符合 `office-kit.windows-live-evidence.v1` 的 JSON 路径并手动触发工作流：

```bash
gh workflow run windows-office-live.yml \
  -f ref=v0.6.0 \
  -f evidence_path='C:\OfficeKit\evidence\windows-office-live.json'
```

`scripts/validate-windows-live-evidence.mjs` 会 fail closed 检查 Windows 平台、Excel/PowerPoint 安装与版本、观察日期、提交 SHA、两个 live workflow 的 `passed` 结果，以及每个应用的 manifest 上传、配对、未保存读写、显式保存、断开重连、源保护和 bridge 空闲退出；PowerPoint 还必须提供双演示文稿隔离、当前选区读取、单页图像复核和不支持能力拒绝。它拒绝 `mock`/`macos` 来源。证据必须由人工在 Windows Office 主机观察产生；签名、截图、录屏和详细操作日志仍由发布负责人按组织流程保存。

## 记录格式

每次 slow/Windows 运行都在交接或 release 文档中记录：触发原因、commit、workflow URL、结论、跳过项及其环境原因。`npm test` 通过只代表 fast gate 通过，不得写成完整 Office/PDF fidelity 或 Windows native acceptance。
