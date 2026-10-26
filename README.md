# opening-report-writer | 开题报告写作与格式

**Research Proposal Writing Skill for Chinese Academic Proposals**

从“只有题目，不知道每节写什么”到有依据的开题报告草稿：学习参考报告的结构和格式，分析题目，规划章节，组织研究现状，完成用户授权的写作与Word整理。

这是给支持Skill的AI助手使用的指令包，不是独立运行的软件，也不自带模型或论文数据库。不默认本科或研究生层次，不限定病理、深度学习或固定研究方向。

采用 [MIT License](LICENSE)。本项目只包含通用写作指令和虚构教学案例，不附带私人报告、学校模板或论文全文。

## 能做什么

- 根据题目和材料确定范围、段落职责、方法与评价关系，不默认学位层次。
- 从正式模板、范文与截图中区分可核实的格式信息，说明共性、差异和适用范围，不把旧范文当成当前学校规定。
- 分开处理写法学习、规划、诊断、撰写和仅排版；学习或审阅不自动改文件。
- 按句子功能处理混合段落，保留有依据的进展和必要课题定位，移出实施细节后检查论证是否完整。
- 区分缺文献事实、缺比较、缺衔接和措辞问题；已有证据充分时直接说明进展，不不断堆风险提示。
- 默认采用五部分骨架，优先服从用户的当前正式模板；保留可编辑公式、图表引用和渲染检查要求。

这不是论文数据库、查重工具、研究成果验证器或自动执行实验的程序。不能保证创新成立、学校验收、发表或方法效果。其他领域可以使用；医学相关提醒仅在对应课题中启用。

## 项目结构

```text
opening-report-writer/
  README.md
  LICENSE
  opening-report-writer/
    SKILL.md
    LICENSE
    agents/openai.yaml
    references/
      topic-to-plan.md
      section-writing-guide.md
      research-status-cases.md
      format-rules.md
```

入口是 [SKILL.md](opening-report-writer/SKILL.md)。[章节指南](opening-report-writer/references/section-writing-guide.md)包含研究现状的判断与修复规则；[案例](opening-report-writer/references/research-status-cases.md)全部为虚构教学材料；[格式指南](opening-report-writer/references/format-rules.md)提供暂定设置，不冒称学校标准。

## 安装与调用

将内层 `opening-report-writer` 整个目录安装到使用环境认可的Skill目录，不能只复制SKILL.md，也不要把仓库外层名称当作Skill名称。

Codex手动安装位置（2026-10-26核对官方Agent Skills文档）：

- 仅供某个项目使用：`<项目根目录>/.agents/skills/opening-report-writer/`。
- 个人使用：`~/.agents/skills/opening-report-writer/`；Windows的 `~` 指用户主目录。

安装后应能找到 `<安装位置>/opening-report-writer/SKILL.md` 及其 `agents/`、`references/`。这是相对结构说明，不依赖维护者的电脑。部分随环境安装器仍使用 `$CODEX_HOME/skills`（未设置时为 `~/.codex/skills`）；按实际客户端和安装器说明选一个位置，不在多个位置重复安装同名Skill。

在已解压的本项目根目录中，可用下面的PowerShell命令进行个人安装。目标已存在时会停止，不覆盖旧版：

```powershell
$source = Join-Path (Get-Location) 'opening-report-writer'
$skills = Join-Path $HOME '.agents/skills'
$target = Join-Path $skills 'opening-report-writer'
if (-not (Test-Path -LiteralPath (Join-Path $source 'SKILL.md'))) {
    throw 'Run this from the extracted repository root.'
}
if (Test-Path -LiteralPath $target) {
    throw 'A skill already exists at the destination; review it before replacing.'
}
New-Item -ItemType Directory -Path $skills -Force | Out-Null
Copy-Item -LiteralPath $source -Destination $target -Recurse
```

Codex会检测Skill变化；没有显示时重启客户端。手动安装路径依据官方文档，安装器路径依据本环境附带说明；这不等于已完成其他客户端的兼容性测试。

手动安装时，先检查目标位置：若已有同名Skill，不直接覆盖，应备份或选用隔离环境。复制后确认SKILL.md旁仍有agents和references目录，再在支持该Skill的环境中显式调用 `$opening-report-writer`。自动选择保持开启，但真实触发取决于客户端与请求，本包未进行跨客户端自动触发测试。

若环境提供 `skill-installer`，可直接使用以下请求：

```text
请使用 skill-installer，从 GitHub 仓库 henkuaia/opening-report-writer
安装 opening-report-writer 子目录。已有同名 Skill 时先说明，不自动覆盖。
```

## 输入与依赖

核心是Markdown指令和案例，不包含可执行脚本，不要求安装其他学术写作Skill。需要能加载Skill的助手及完成相应任务所需的用户材料。题目可以先用于暂定规划；事实性综述还需要实际文献证据。

| 任务 | 所需能力 | 缺少时如何处理 |
| --- | --- | --- |
| 题目分析与结构规划 | 读取题目和用户材料 | 明确暂定范围，不补造事实 |
| 真实论文学习与综述 | 可读原文；需要补证时有检索或获取渠道 | 列出具体证据缺口，不以案例冒充真实文献 |
| DOCX编辑与交付 | 可用的文档生成或编辑工具 | 文本分析仍可继续；不能称已经生成Word |
| 最终版式检查与PDF | DOCX渲染或导出工具及页面查看能力 | 明确未渲染或未逐页查看，不宣称版式通过 |

`documents` 是环境提供时优先使用的文档工作流，不是本包捆绑依赖。`pathology-paper-writing` 等领域Skill只是相关任务中的可选增强，不是通用写作的前置条件。本包不宣称它们存在统一公开安装渠道；缺少时不阻塞文字层面的任务。精确模板复刻须用户另外提供有权使用的模板。

## 调用示例

**从题目开始：** “请使用 $opening-report-writer。题目是‘面向低功耗无线传感节点的自适应采样方法研究’，目前没有确定算法或实验数据。请给暂定章节提纲，逐节说明要写什么、需要什么证据；不要编造论文或替我确定最终方法。”

**学习参考报告：** “请使用 $opening-report-writer，学习我提供的两份参考报告和当前正式模板。分别总结章节职责、段落推进和格式规则，给来源位置；说明差异、可迁移内容与未知项。本轮只总结，不修改报告或Skill。”

**只学习：** “请使用 $opening-report-writer，阅读我提供的论文引言与相关工作，说明研究现状如何组织，标出实际版本与位置；只总结，不修改报告。”

**诊断：** “请使用 $opening-report-writer，诊断所附研究现状。按句子功能说明保留、压缩或移位，检查移位后缺少的比较与证据；不要改文件。”

**授权修改：** “请使用 $opening-report-writer，根据我提供的报告和已核实文献，只修改研究现状。保留成立的内容，不确定的方案不要代定；另存新版并说明验证范围。”

**格式：** “请使用 $opening-report-writer，按所附正式模板整理报告，保留正文、数学公式和引用；另存DOCX，有工具时从最终DOCX导出PDF并逐页检查。”

## 来源与公开处理

本项目将开题报告辅助写作过程中形成的通用规则整理为独立Skill；写法案例为虚构教学材料，不构成真实文献证据。未捆绑其他Skill的实现、第三方论文正文、学校模板、报告、图片、项目数据、身份信息或历史对话。

用户自行提供的论文、模板和报告不因使用本Skill而适用本项目的MIT许可；使用者仍需遵守这些材料各自的使用条件。核心写作不依赖维护者的本机目录或私人文件。

## 许可

Copyright (c) 2026 henkuaia。完整条款见 [LICENSE](LICENSE)。Skill子目录包含相同的许可文件，便于单独安装与再分发时保留版权和许可声明。

## 验证范围

结构检查使用当前环境的skill-creator验证器，并额外检查相对链接、UI元信息、文件类型和不应公开的路径内容。隔离检查仅将内层Skill复制到临时目录，解析入口并读取其全部引用，不改当前安装位置。

四个行为案例覆盖混合段落、合理段落、非病理题目以及授权和工具限制。判断的是事实保留、比较维度、章节职责与诚实交付，不只匹配关键词；输入、处理示范和判断依据见[写法案例](opening-report-writer/references/research-status-cases.md)。这些是整理时的会话内自检，不是独立评测或跨模型测试。

格式指南另附虚构的“当前要求与旧范文冲突”案例，用于检查模板优先级与文本读取边界，不冒充真实Word排版测试。安装命令仅在临时主目录中验证复制结构和已有目标保护，不代表已在其他电脑安装。

文件检查不能证明自动发现、文献准确性或写作质量；当前会话行为自检不是独立评测。没有在其他电脑安装，也没有生成报告来重测DOCX、公式或排版能力。格式指南的保留不等于这些能力已经端到端验证。
