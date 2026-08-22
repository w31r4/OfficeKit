# OfficeKit 门禁分层

OfficeKit 把“开发反馈速度”和“发布可信度”分开。小改动先走 fast gate；会下载外部 provider、启动 LibreOffice/Poppler、运行 PromptBench、渲染完整 reference Skill 矩阵、重跑全部模板或重建 WASM 的检查，只在 slow gate、发布候选或 nightly 执行。

## 三类入口

| 层 | 本地入口 | Hosted 入口 | 触发 | 包含 |
| --- | --- | --- | --- | --- |
| Fast | `npm test`（等同 `npm run test:fast`） | `.github/workflows/ci.yml` | 每次 push/PR | JS syntax/import、核心四格式模型、OfficeKit 路由/参考插件路径同步/可移植性 validator、Help、包内容 smoke。不会下载 Python/PyMuPDF/qpdf/OCR/veraPDF，不运行完整域 Skill 或 reference Skill render/matrix、PromptBench、clean-install 或 20 套默认模板的全量原生回归。 |
| Slow | `npm run test:slow` | `.github/workflows/ci-slow.yml` | 每日定时或人工冻结 milestone；不随普通 push/PR 自动触发 | 原完整测试链、provider/pack、Playwright/LibreOffice/Poppler、PromptBench candidate/reference、默认模板全量 import/export/recalc/render、examples、release/package、OfficeBridge 与 OfficeKit WASM。Hosted workflow 将同一条步骤表按十个职责段顺序执行，避免单个长命令被 runner 取消。 |
| Windows Office live | 见下文 | `.github/workflows/windows-office-live.yml` | 手动排队；Live host 变更或 release candidate | 真实 Windows + Microsoft Office 人工观察证据。GitHub-hosted Windows、macOS mock、Add-in build smoke 和 CLI/package smoke 都不能替代这条证据。 |
| Windows PPTX lossless | 见下文 | `.github/workflows/windows-pptx-lossless.yml` | 三份复杂 PPTX 的无损 Goal 验收 | 真实 Windows PowerPoint 对三份冻结样本执行打开、浏览、局部编辑、保存副本、重新打开、非目标页像素比较和不支持能力拒绝；只接受人工证据，不把 Live Add-in 或 macOS 结果代替它。 |

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

## Windows PPTX 无损编辑证据

三份复杂样本的独立验收使用 `.github/workflows/windows-pptx-lossless.yml`，而不是上面的 Live Add-in lane。操作员先在真实 Windows x64 PowerPoint 中按 `evals/pptx-lossless/manifest.v1.json` 的源 SHA 打开每份文件，完成声明的局部编辑，保存到不同路径并重新打开；随后用 PowerPoint 产生非目标页像素比较、修复提示和高级对象保全记录，再写入 `office-kit.windows-pptx-lossless-evidence.v1` JSON。`visualReview.pageComparisons` 必须覆盖三份样本的全部 48 页，逐页绑定源/输出像素 SHA-256；目标页必须有像素变化，非目标页必须是相同指纹。校验器
`scripts/validate-windows-pptx-lossless-evidence.mjs` 会绑定当前 checkout SHA、三份冻结源 SHA、目标节点、保存副本、重新打开、源保护、非目标页像素一致和不支持能力拒绝；它拒绝 macOS、mock、LibreOffice-only 或缺少任一源的证据。

```bash
gh workflow run windows-pptx-lossless.yml \
  -f ref=09cd0723ae9d150af08f34b5bafdad20776f1b42 \
  -f evidence_path='C:\\OfficeKit\\evidence\\windows-pptx-lossless.json'
```

为减少手工拼接证据，Windows 操作员可以先把三份 OfficeKit 输出按
`<source-id>.pptx` 放入输出目录，再运行
`scripts/collect-windows-pptx-lossless-evidence.ps1`。该脚本通过真实
PowerPoint COM 打开源文件和输出，保存并重新打开独立副本，导出全部页面
PNG、计算逐页 SHA-256，并逐项询问无法由 COM 自动证明的人工观察结果：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File `
  scripts/collect-windows-pptx-lossless-evidence.ps1 `
  -SourceRoot C:\OfficeKit\assets `
  -OutputRoot C:\OfficeKit\outputs `
  -EvidencePath C:\OfficeKit\evidence\windows-pptx-lossless.json `
  -Commit 09cd0723ae9d150af08f34b5bafdad20776f1b42
```

脚本不会把人工确认项默认填成通过；任一确认失败都会拒绝写出完整证据。
输出目录中的三个文件名必须是 `suanzhi-future-2026.pptx`、
`blue-gray-acid-template.pptx` 和 `mckinsey-customer-loyalty.pptx`。

这条 lane 只验证人工已经观察并记录的结果，不会把“workflow 通过”误写成 PowerPoint 已验收；没有可用的 self-hosted Windows Office runner 时应保持未完成。

## 记录格式

每次 slow/Windows 运行都在交接或 release 文档中记录：触发原因、commit、workflow URL、结论、跳过项及其环境原因。`npm test` 通过只代表 fast gate 通过，不得写成完整 Office/PDF fidelity 或 Windows native acceptance。
