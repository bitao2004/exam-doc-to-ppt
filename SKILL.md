---
name: exam-doc-to-ppt
description: 将考试题目文档（.docx，支持单选/多选/判断等可解析题型）转为 PowerPoint 幻灯片，支持"每页 N 题 + 答案放映时隐藏、点击鼠标逐题揭示"的刷题模式。当目标播放器包含 WPS 时，必须使用纯幻灯片切换实现点击显隐（WPS 会静默丢弃手写动画时间线）。This skill should be used when the user wants to turn an exam/question bank docx into a click-to-reveal slideshow, or asks to hide answers until clicked in a PPT. 默认输出即精致美化版（渐变页眉、卡片柔影、彩色题号徽标、正确答案绿色高亮+✓）。
agent_created: true
---

# 考题文档转幻灯片（Exam Doc → PPT）

## Overview

把"考题/题库 Word 文档"转成可放映的 PowerPoint，用于刷题或课堂讲评。核心诉求通常是：**答案先藏起来，点击鼠标再显示**。本技能把经过多轮踩坑验证的可靠方案固化下来，重点解决 WPS 兼容性问题。脚本默认输出即为**精致美化版**（无需额外开关）。

## 关键经验（必读，决定成败）

> **WPS 会静默丢弃手写的动画时间线（`p:timing`）。** 不要用 python-pptx 手写"出现"入场动画或 `p:set` 改 `style.visibility` 来做点击显隐——PowerPoint 可能认，但 WPS 一旦遇到不合规范的 XML 就整块丢弃，结果答案全程直接显示。社区（python-pptx #1106、AWS Builder 文章）已确认此行为。

**可靠方案：用「幻灯片切换」实现点击显隐，零动画。**
把"点击显示答案"拆成多张物理幻灯片，每张用静态文本表达一种揭示状态，放映时点鼠标自动前进。WPS / PowerPoint 100% 支持，永不失效。

## 工作流

### 1. 解析题目
从 docx 读取段落，按行特征切分题目。通用正则（微调以适配实际格式）：
- 题号行：`^\s*(\d+)\s*[\.、]\s*(.*)`
- 选项行：`^\s*([A-Fa-f])[\.、]\s*(.*)`（已覆盖多选的 A–E 共 5 选项）
- 答案行：`^\s*参考*答案\s*[:：]\s*([A-F]+)`（多选答案可为多字母，如 `BCE`/`ABCDE`，解析时去空格）

**题型自适应**：单选 / 多选 / 判断 通用。判断通常是 `A. 正确 / B. 错误`（2 选项），直接复用即可；**多选题答案含多个字母时，所有正确选项行都会绿色高亮并显示 ✓**（脚本用 `letter in answer` 判断，而非单字母相等）。页眉标题按**文件名关键字 → 正文关键字**自动识别（见步骤 3）。

> ⚠️ **不要按"文档首段"识别标题**：考题文档的首段永远是题号行（如 `1. 题干…`），据此判断必然回退成"单选题"。历史上这里踩过坑——早期实现只看首段，导致多选/判断题的页眉全部错标为"单选题"，而验证时又被"扫描全文命中题型词"的假阳性掩盖。现改为文件名/正文关键字优先。

用 `python-docx` 读取；先 `print` 出前若干非空段落确认格式，再写完整解析。校验：题号是否 1..N 连续、每题选项数是否为 4、答案是否齐全。

### 2. 选择"揭示模式"
设每页 `per_page` 题（默认 2）。每组生成 `k` 张物理幻灯片表达逐步揭示：

| 模式 | 物理页数/组 | 行为 | 适用 |
|---|---|---|---|
| `group`（整组） | 2 | A 全隐 → B 全显 | 点一下两题答案一起出 |
| `per-question`（逐题，默认） | `per_page + 1` | 第 0 张全隐 → 每多一张多显 1 题 → 末张全显 | 分别点击、逐题揭示 |

例：每页 2 题 + `per-question` → 每组 3 张：A`[？,？]` → B`[A,？]` → C`[A,B]`。点击前进即逐题显示。

### 3. 生成幻灯片（运行脚本）
直接用 `scripts/generate_ppt.py`（已参数化）：

```bash
python generate_ppt.py --src 考题.docx --out 考题.pptx \
    --per-page 2 --reveal per-question --layout horizontal
```
- `--title`：页眉标题。识别优先级 **`--title` 显式指定 > 源文件名含题型关键字（多选/判断/单选）> 正文关键字 > 默认"单选题"**；一般无需指定。
- `--layout vertical`（默认）：每页 `per_page` 题**上下堆叠**，单卡宽扁。
- `--layout horizontal`：每页 `per_page` 题**左右双栏**并排，单卡接近方形、充分利用宽屏空间（要求 `per_page>1`；`per_page=1` 时自动退化为单卡满版）。

依赖（在隔离 venv 安装，勿污染全局）：
```bash
python -m venv <venv> && <venv>/Scripts/python.exe -m pip install python-pptx python-docx lxml
```

脚本默认即**精致美化版**，视觉元素（见 `generate_ppt.py` 顶部常量，均可改）：
- **柔和浅色渐变背景**（`BG`→`BG2`，极浅冷白→浅雾蓝，**非深色纯色**，避免刺眼）+ **柔和蓝紫渐变页眉**（`HDR_C1`→`HDR_C2`，比深靛蓝更轻盈）；右侧白色圆角页码药丸「第 N / M 组」。**页眉不再显示"点击鼠标显示答案"等副标题文字**，画面更干净。
- **白色圆角卡片**：细边框 + 柔和投影（`outerShdw`）+ 左侧蓝色彩脊。
- **题号徽标**：蓝色渐变小圆 + 白色题号居中；题干首行与题号圆**顶部对齐**。
- **选项行**：浅灰圆角胶囊、左侧蓝色圆形字母标；**正确答案行绿色高亮**（绿底 + 绿边框 + 绿字母圆 + ✓ 绿勾），一眼定位；选项行距收紧、与末选项到答案行之间留有间距。
- **页脚**：左侧提示药丸（👉 点击显示答案 / ✅ 答案已显示）+ 右侧右对齐「参考答案：X / ？」。
- **长文本自动缩小字号**：题干/选项文字过长时，`fit_font()` 按"宽×高"估算自动降到能完整装进区域的最小字号（题干下限 10.5pt、选项下限 9pt），**保证每道题目（含题干与选项）完整显示、不溢出卡片**；短文本仍用默认大字号。
- 所有文本按卡片**宽×高双维度比例布局**，答案/提示始终收在卡片内、**不溢出**。

### 3.1 美化与主题
- 开关：`--no-beautify` 关闭装饰（渐变/柔影/彩脊/徽标），退回朴素实心版；默认开启。
- 改主题：直接改 `generate_ppt.py` 顶部 `HDR_C1/HDR_C2`(页眉)、`BADGE_*`(题号)、`GREEN_*`(答案高亮)、`LETTER_*`(字母标) 等 hex 常量即可换配色。
- 字号随每页题数自适应：每页 ≤2 题用大字号（与精致版一致）；>2 题自动压缩并提示拥挤。
- **左右双栏（`--layout horizontal`）**：`build_card` 改为按卡片「宽×高」双维度比例布局——单卡接近方形时自动放大字号、选项行更舒展，并让两题并排充分利用宽屏。校验时务必确认：两卡片 x 区间不重叠、各自文本收在所属卡片内无溢出。

### 4. 校验（每次必做）
重新用 python-pptx 打开产物，断言：
- 物理幻灯片数 == 组数 × (group?2:(per_page+1))
- 每组状态页答案文本符合预期（A 全 `？`、末页全为真实字母）
避免"声称成功却打不开"。

### 5. 文件锁定坑
若目标 `.pptx` 正被预览面板/WPS 打开，写入会 `Permission denied`。处理方法：
- 先尝试原文件名；失败则写入备用名（如 `考题_分别点击.pptx`）并提示用户关闭旧预览后改名/覆盖。
- 切勿静默丢结果。

## 推荐参数（开箱即用）
针对"单选题题库 + 宽屏放映 + 逐题点击揭示"的典型场景，直接用：
```bash
python generate_ppt.py --src 单选题.docx --out 单选题.pptx \
    --per-page 2 --reveal per-question --layout horizontal
```
- `--per-page 2` + `--reveal per-question`：每屏 2 题、分别点击逐题揭示（每组 3 张物理幻灯片）。
- `--layout horizontal`：左右双栏，充分利用 16:9 宽屏；`per_page=1` 时自动退化单卡满版。
- 仅要朴素版时加 `--no-beautify`。

## 常见失败模式速查表（踩坑汇总）
| 现象 | 根因 | 对策 |
|---|---|---|
| WPS 放映时答案直接显示 | 手写动画 `p:timing` 被 WPS 静默整块丢弃 | 一律用「幻灯片切换」揭示，零动画（本技能默认方案，永不失效） |
| 目标 `.pptx` 写入 `Permission denied` | 文件被预览面板/WPS 独占 | 脚本自动改存备用名并**打印警告**；关闭预览后重命名覆盖 |
| 长题干/长选项文字被截断、溢出卡片 | 固定字号装不下 | `fit_font()` 按"宽×高"自动缩到可完整显示的最小字号（题下限 10.5pt / 选项下限 9pt） |
| 页眉出现多余提示文字 | 副标题未管控 | `SHOW_SUBTITLE=False`（默认）已移除"点击鼠标显示答案"等文字；恢复改 `True` |
| 深色纯色背景刺眼 | 配色过硬 | 已改柔和浅色渐变（BG→BG2）+ 柔和蓝紫渐变页眉；换色改脚本顶部 hex 常量 |
| 左右双栏下两卡重叠/文本溢出 | 单卡宽高比例算错 | `build_group(horizontal)` 按 `SLIDE_W_IN` 均分；校验断言 x 区间不重叠、文本收在卡内 |

> 用户明确偏好（已固化进默认行为）：① 答案必须点击后才显示，不依赖 WPS 不认的动画；② 长题必须完整显示（自动缩字号）；③ 页眉去掉"点击鼠标显示答案"等提示文字；④ 背景走柔和浅色渐变，不用深色纯色。

## 技能维护（重打包）
迭代本技能后，重新打包为可分发的 zip：
```bash
python "<skill-creator>/scripts/package_skill.py" \
    --skill-dir "C:/Users/guosp/.workbuddy/skills/exam-doc-to-ppt" \
    --out "C:/Users/guosp/.workbuddy/skills/exam-doc-to-ppt/exam-doc-to-ppt.zip"
```
`<skill-creator>` 即 skill-skill-creator 插件目录，例如
`C:/Users/guosp/.workbuddy/plugins/cache/workbuddy-builtin/skill-skill-creator/0.1.0`。
打包会包含 SKILL.md 与 scripts/。改主题/布局只需改 `generate_ppt.py` 顶部常量，无需动结构。

## References
- 详细 OOXML 动画兼容性背景见社区资料（python-pptx GitHub #1106）；结论已上述：**WPS 下不依赖动画，用翻页**。
- 当前未单独建 references 文件——核心知识已在本文；如后续需扩展题型解析正则模板，再补 `references/parsing.md`。

## Resources
- `scripts/generate_ppt.py`：完整、参数化、可直接跑的生成脚本（含解析、两种揭示模式、美化布局、`fit_font` 自适应字号、左右双栏）。
- `exam-doc-to-ppt.zip`：本技能的分发包（由 `package_skill.py` 重打包生成，含 SKILL.md + scripts/）。
- `README.md` / `LICENSE` / `.gitignore`：面向 GitHub 分享的配套文件（`README.md` 面向人类读者，`SKILL.md` 面向 WorkBuddy 路由，两者定位不同，不要合并）。
- 踩坑速查见上文「常见失败模式速查表」；解析正则模板如需扩展题型再补 `references/parsing.md`。
