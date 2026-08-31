# exam-doc-to-ppt · 考题文档转幻灯片

把考试题目 Word 文档（`.docx`）转成可放映的 PowerPoint，用于刷题或课堂讲评。

核心能力：**答案先隐藏，点击鼠标逐题显示** —— 且在 **WPS 里也稳如老狗**。

---

## 为什么需要它

"点击显示答案"最直觉的做法是给 PPT 加"出现"入场动画。但在 WPS 里这条路走不通：

> **WPS 会静默丢弃手写的动画时间线（`p:timing`）。**
> 用 python-pptx 手写入场动画或 `p:set` 改 `style.visibility`，PowerPoint 可能认，
> WPS 一旦遇到不合规范的 XML 就整块丢弃 —— 结果答案全程直接显示，无法点击揭示。

本技能改用**纯「幻灯片切换」**实现：把揭示过程拆成多张物理幻灯片，每张用静态文本表达一种状态，放映时点鼠标前进。零动画，WPS / PowerPoint 100% 支持，永不失效。

例：每页 2 题、逐题揭示 → 每组 3 张物理幻灯片

```
第 1 张： [？, ？]   ← 答案全隐藏
第 2 张： [A , ？]   ← 点一下，显示第 1 题答案
第 3 张： [A , B ]   ← 再点一下，显示第 2 题答案
```

---

## 特性

- **多题型通吃**：单选 / 多选（A–E 五选项）/ 判断（`A.正确 / B.错误`）
  - 多选题答案含多个字母（如 `BCE`、`ABCDE`）时，**所有正确选项行都会绿色高亮并打 ✓**
  - 页眉标题自动识别（文件名关键字 → 正文关键字 → 默认"单选题"）
- **两种揭示模式**：`per-question` 逐题揭示（默认）/ `group` 整组一次揭示
- **两种版式**：`horizontal` 左右双栏（16:9 宽屏推荐）/ `vertical` 上下堆叠
- **默认即精致美化**：柔和浅色渐变背景、蓝紫渐变页眉、白色圆角卡片 + 柔影 + 彩脊、题号徽标、答案绿色高亮 + ✓、页码药丸
- **长文本自动缩字号**：`fit_font()` 按「宽 × 高」估算，题干降到可完整显示为止（下限 10.5pt / 选项 9pt），**绝不截断**
- **自适应防溢出**：所有文本按卡片宽高双维度比例布局，答案与提示始终收在卡片内
- 主题配色全部集中在脚本顶部 hex 常量，改色不用动结构

---

## 目录结构

```
exam-doc-to-ppt/
├── SKILL.md              # WorkBuddy 技能说明（frontmatter 合规，可直接被识别）
├── README.md             # 本文件
├── LICENSE               # MIT
├── .gitignore
└── scripts/
    └── generate_ppt.py   # 生成脚本本体（唯一依赖代码）
```

> `exam-doc-to-ppt.zip` 是打包分发产物，由 `package_skill.py` 从源码生成，已在 `.gitignore` 中排除。

---

## 环境依赖

- Python 3.9+
- `python-pptx`、`python-docx`、`lxml`

建议在虚拟环境中安装，不污染全局：

```bash
python -m venv .venv
# Windows
.venv\Scripts\python.exe -m pip install python-pptx python-docx lxml
# macOS / Linux
.venv/bin/python -m pip install python-pptx python-docx lxml
```

---

## 快速开始

```bash
python scripts/generate_ppt.py \
    --src 单选题.docx \
    --out 单选题.pptx \
    --per-page 2 \
    --reveal per-question \
    --layout horizontal
```

打开生成的 `.pptx`，按 F5 放映，点鼠标即可逐题显示答案。

### 参数

| 参数 | 默认值 | 说明 |
|---|---|---|
| `--src` | `考题.docx` | 输入 docx 路径 |
| `--out` | `考题.pptx` | 输出 pptx 路径 |
| `--title` | 自动识别 | 页眉标题；识别优先级：显式指定 > 文件名关键字 > 正文关键字 > `单选题` |
| `--per-page` | `2` | 每页题数 |
| `--reveal` | `per-question` | `per-question` 逐题揭示 / `group` 整组揭示 |
| `--layout` | `vertical` | `vertical` 上下堆叠 / `horizontal` 左右双栏 |
| `--no-beautify` | — | 关闭美化，退回朴素实心版（默认开启美化） |

### 输出页数

```
物理幻灯片数 = 组数 × (reveal == group ? 2 : per_page + 1)
组数         = ceil(题目总数 / per_page)
```

例：304 题、`--per-page 2 --reveal per-question` → 152 组 × 3 = **456 张**。

---

## 输入格式要求

按段落解析，只要符合下列特征即可（支持中文标点）：

```
1. 题干文字……
A. 选项一
B. 选项二
C. 选项三
D. 选项四
参考答案：C
```

对应正则：

| 行类型 | 正则 |
|---|---|
| 题号行 | `^\s*(\d+)\s*[\.、]\s*(.*)` |
| 选项行 | `^\s*([A-Fa-f])[\.、]\s*(.*)` |
| 答案行 | `^\s*参考*答案\s*[:：]\s*([A-F]+)`（多选可为 `BCE`，自动去空格） |

题目之间可以有空行，顺序需为「题号 → 选项 → 答案」。遇到新题号或答案行即切分下一题。

> **建议**：换用新来源的文档时，先 `print` 出前若干非空段落确认格式，再跑全量生成。

---

## 自定义主题

打开 `scripts/generate_ppt.py`，改顶部的 hex 常量即可：

| 常量 | 作用 |
|---|---|
| `BG` / `BG2` | 背景渐变（默认极浅冷白 → 浅雾蓝，**非深色**，不刺眼） |
| `HDR_C1` / `HDR_C2` | 页眉渐变（默认柔和蓝紫） |
| `BADGE_*` | 题号徽标配色 |
| `GREEN_*` | 正确答案高亮配色 |
| `LETTER_*` | 选项字母标配色 |

另有 `SHOW_SUBTITLE`（默认 `False`）可控制是否显示页眉副标题文字。

---

## 作为 WorkBuddy 技能安装

把整个目录复制到用户级技能目录即可被自动识别：

```bash
# Windows
cp -r exam-doc-to-ppt ~/.workbuddy/skills/
# 或作为某仓库的子目录
# <repo>/skills/exam-doc-to-ppt/
```

`SKILL.md` 的 frontmatter 已包含 `name` / `description`，合规可路由。

---

## 已知坑

| 现象 | 原因 / 对策 |
|---|---|
| WPS 放映时答案直接显示 | 手写动画被 WPS 丢弃。本技能不用动画，用幻灯片切换，不受影响 |
| 目标 `.pptx` 写入 `Permission denied` | 文件被预览面板 / WPS 独占。脚本会自动改存备用名并**打印警告**，不会静默丢结果 |
| 长题干被截断 | `fit_font()` 会自动缩字号；若仍溢出，检查是否超出下限（题干 10.5pt） |
| 多选 / 判断题页眉显示成"单选题" | 标题按文件名 / 正文关键字识别，把文件命名为 `多选题.docx` 或正文含题型词即可；也可用 `--title` 显式指定 |

---

## License

[MIT](LICENSE) © 2026 郭绍坡
