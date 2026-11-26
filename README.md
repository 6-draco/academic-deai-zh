# academic-deai-zh

面向中文学术论文的双阶段降 AI 率改写技能（Agent Skill / Antigravity Plugin），附带 `.docx` 无损样式抽取、校验与写回脚本。

在严格保留专业术语、引用格式、数据事实与论证确定性程度的前提下，分两阶段消除大语言模型常见的模板化结构与低困惑度套话，并输出可逐条核查的修改对照表。

## 核心特点

- **五条学术红线**：零事实篡改、锁定核心锚点（术语/引注/数据/直接引语/公式/图表号）、保留确定性程度（Epistemic Modality）、守住学术书面语体、字数严格控制在原文 85%–115%。
- **双阶段改写**：
  1. **第一阶段（De-Templating）**：拆解机械序数框架、虚化动词、多层“的”字长定语、同义叠加四字格、均匀句长、高频被动句与空泛首尾套话。
  2. **第二阶段（Academic Calibration）**：降频高频 AI 学术套词，隐性化句间衔接（主位推进/概念呼应），微调局部因果视角，并逐句回检事实与引注锚定。
- **四部分可核查输出**：锁定锚点清单 → 纯净改写定稿 → 关键改动对照表（Diff）→ 硬性自检清单。
- **`.docx` 安全工作流**：自动跳过标题、作者信息、关键词行、目录、图表题注与参考文献列表；支持硬性锚点/字数自动核查；按原段落样式与普通/加粗 run 写回 `<原文件名>_deai.docx`，严禁覆盖原文件，对含上标引注/公式/超链接等复杂段落默认拦截保护。

## 仓库结构

本仓库采用多技能 / Antigravity Plugin 目录结构，便于后续在同一仓库中扩展更多学术写作技能：

```text
academic-deai-zh/
├── plugin.json                         # Antigravity Plugin 清单
├── README.md                           # 说明文档
├── requirements.txt                    # Python 依赖（处理 .docx 时需要）
├── LICENSE                             # MIT 协议
└── skills/
    └── academic-deai-zh/
        ├── SKILL.md                    # 技能主提示词与工作流规范
        └── scripts/
            └── docx_tool.py            # .docx 段落抽取、核查与无损写回脚本
```

## 安装方式

### 方式一：作为 Antigravity Plugin 安装（推荐）

将仓库克隆到全局插件目录（所有项目生效）：

```bash
git clone https://github.com/6-draco/academic-deai-zh.git ~/.gemini/config/plugins/academic-deai-zh
```

### 方式二：作为独立 Skill 安装

仅安装 `skills/academic-deai-zh` 目录：

- **全局生效**：复制到 `~/.gemini/config/skills/academic-deai-zh`
- **仅当前论文项目生效**：复制到项目根目录 `.agents/skills/academic-deai-zh`

### 方式三：作为项目常驻 Rule 使用

若工作区专用于论文写作，可去掉 `skills/academic-deai-zh/SKILL.md` 顶部的 YAML frontmatter，保存至项目的 `.agents/rules/academic-deai-zh.md`（或全局 `~/.gemini/config/rules/`）。

### 安装 `.docx` 脚本依赖（处理 Word 文档时）

```bash
pip install -r requirements.txt
```

## 调用示例

在对话中直接触发：

- `使用 academic-deai-zh 改写 @ch2.md 的第 3 至第 6 段`
- `用 academic-deai-zh 处理 @paper.docx 第三章正文，确认后生成 _deai 文件`
- 直接粘贴论文段落并注明“使用 academic-deai-zh 降 AI 率改写”

### `docx_tool.py` 命令行用法

```bash
# 1. 抽取段落与跳过/加粗/复杂格式标记
python skills/academic-deai-zh/scripts/docx_tool.py extract paper.docx paras.json

# 2. 校验改写后的引注、数字、英文术语与 85%–115% 字数区间
python skills/academic-deai-zh/scripts/docx_tool.py check edits.json

# 3. 保留样式写回新文件（不覆盖原文件）
python skills/academic-deai-zh/scripts/docx_tool.py apply paper.docx edits.json paper_deai.docx
```

## License

[MIT](./LICENSE)
