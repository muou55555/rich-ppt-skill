---
name: rich-ppt-skill
description: >-
  基于 GordenPPTSkill 扩展的增强版 PPT 生成技能。在保留原版模板替换能力的基础上，新增「富样式原创模式」：
  通过 python-pptx 代码生成带图标（Phosphor Icons 扁平填充风格）、圆角卡片、箭头流程图、高亮表格、
  标注框等精美组件的可编辑 .pptx 文件。支持两种主模式：
  ① 通用PPT（任意主题、商务/技术/学术）；
  ② 技术汇报（架构图、流程、代码注释、对比表格）。
  触发词：做PPT / 生成幻灯片 / 技术汇报 / 架构汇报 / ppt / slides / 演示文稿
---

# rich-ppt-skill

> **版本** `1.0.0` | 基于 GordenPPTSkill v1.x 扩展，新增 Rich Mode（富样式原创生成）

---

## 🚀 启动时第一步：安装依赖 & 克隆基础模板库

```bash
# 1. 安装 Python 依赖
pip install python-pptx cairosvg --break-system-packages -q

# 2. 安装 Phosphor Icons（扁平图标，MIT 开源，与 iconfont 面性风格一致）
npm install --prefix /tmp/phosphor @phosphor-icons/core 2>/dev/null || true

# 3. 克隆 GordenPPTSkill 模板库（获取 19 套内置模板）
git clone --depth 1 https://github.com/GordenSun/GordenPPTSkill.git /tmp/GordenPPTSkill 2>/dev/null || true

# 4. 克隆本 Skill 到可写目录
git clone --depth 1 https://github.com/YOUR_GITHUB_USERNAME/rich-ppt-skill.git /tmp/rich-ppt-skill 2>/dev/null || true
```

> ✅ 同一会话内只需执行一次。

---

## 何时使用本 Skill

- 用户需要生成/制作/创建 PPT、演示文稿、幻灯片
- 用户有 Markdown 文档、技术文章、架构说明，希望转成 PPT
- 用户想要"精美""有图标""专业感"的 PPT
- 技术汇报场景：架构图解、流程说明、对比分析

---

## 三种模式选择

### 判断流程

```
用户需求
  ├─ 提供了 .pptx 模板 ──────────────────→ 模式 B（模板替换，沿用 GordenPPTSkill 流程）
  ├─ 要求"极简/干净/不要花哨" ───────────→ 模式 A（GordenPPTSkill 内置模板）
  └─ 其他（有内容、要精美、技术汇报等）──→ 模式 C（Rich Mode，本 Skill 核心）
```

### 模式 A：GordenPPTSkill 内置模板（文字替换）

沿用 `/tmp/GordenPPTSkill/SKILL.md` 的完整流程，不赘述。

### 模式 B：用户自带模板（文字替换）

沿用 GordenPPTSkill 模式 B 流程。

### 模式 C：Rich Mode（富样式原创生成）★ 本 Skill 核心

**适合：** 技术文档 → PPT、内容信息量大、需要图标卡片/流程图等可视化组件

---

## 模式 C 完整工作流

### Step 1：解析内容，规划幻灯片结构

读取用户提供的文档/大纲，按以下结构规划每页：

```
封面页（D-cover）
目录页（D-toc）
章节内容页（D-section-N）× N
  ├─ 概念/哲学 → 圆角卡片布局（2×2 或 3×2）
  ├─ 流程/步骤 → 箭头流程图（chevron_flow）
  ├─ 对比/清单 → 表格（htable）
  ├─ 时序/架构 → 时序行布局
  └─ 要点说明 → 标注框（callout_box）
结尾/参考页（D-end）
```

### Step 2：生成 build script

调用 `/tmp/rich-ppt-skill/scripts/build_pptx.py` 的核心函数库，
按规划的结构生成 Python 脚本，例如：

```python
from scripts.components import *
prs = new_presentation()
D_cover(prs, title="Claude Code 实现原理", subtitle="架构·机制·复刻路线")
D_toc(prs, chapters=[...])
D_section(prs, title="设计哲学", layout="cards_2x2", items=[
    CardItem(icon="lightbulb", color=NAVY, title="单一主循环", body="..."),
    CardItem(icon="agent",     color=SKY,  title="模型即调度器", body="..."),
    ...
])
save_pptx(prs, "/tmp/output.pptx")
```

### Step 3：嵌入图标

使用 `/tmp/rich-ppt-skill/scripts/icons.py` 的 `get_icon(name, color, size)` 函数。
图标来源：**Phosphor Icons fill**（MIT 开源，扁平实色，视觉等同 iconfont 面性图标）。

可用图标清单见 `references/icon-reference.md`。

### Step 4：运行生成脚本

```bash
python3 /tmp/rich-ppt-skill/scripts/build_pptx.py \
    --input your_script.py \
    --output /path/to/output.pptx
```

或直接在 Claude 会话中通过 Bash 工具运行生成脚本。

### Step 5：交付

将生成的 .pptx 复制到用户工作目录，调用 `present_files` 展示。

---

## 视觉规范（Rich Mode 配色与排版）

### 色板

```python
NAVY   = RGBColor(0x00, 0x32, 0x9D)  # 主色：深海蓝  #00329D
SKY    = RGBColor(0x44, 0x72, 0xC4)  # 辅色：天空蓝  #4472C4
MED    = RGBColor(0x30, 0x55, 0x98)  # 中蓝          #305598
STEEL  = RGBColor(0x2E, 0x75, 0xB6)  # 钢蓝          #2E75B6
GREEN  = RGBColor(0x1E, 0x8C, 0x55)  # 绿色          #1E8C55
GOLD   = RGBColor(0xED, 0xAD, 0x1A)  # 金色          #EDAD1A
ORANGE = RGBColor(0xE0, 0x6B, 0x1A)  # 橙色          #E06B1A
RED2   = RGBColor(0xC0, 0x39, 0x2B)  # 警示红        #C0392B
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
CARD   = RGBColor(0xEB, 0xF0, 0xF8)  # 卡片背景浅蓝
LINE   = RGBColor(0xCC, 0xD8, 0xEE)  # 分隔线
DARK   = RGBColor(0x1A, 0x1A, 0x2E)  # 正文深色
GRAY   = RGBColor(0x66, 0x72, 0x82)  # 副标题灰
```

### 字体规范

| 元素 | 字号 | 粗细 | 颜色 |
|---|---|---|---|
| 封面主标题 | 32pt | Regular | WHITE/NAVY |
| 页面标题 | 28pt | Regular | NAVY |
| 卡片标题 | 13-14pt | Bold | accent色 |
| 正文 | 10-11pt | Regular | DARK |
| 副标题/说明 | 9-11pt | Regular | GRAY |
| 标签/徽章 | 9-10pt | Bold | WHITE |

### 幻灯片尺寸（宽屏 16:9）

```python
W = 12192000  # EMU，宽
H = 6858000   # EMU，高
```

---

## 组件参考（Rich Mode）

### 标题块 title_block(slide, title, subtitle='')

白底左侧竖条（NAVY）+ 标题文字（NAVY 28pt）+ 水平分隔线 + 灰色副标题。
与 GordenPPTSkill 模板视觉语言对齐。

### 圆角卡片 rcard(slide, x, y, w, h, title, body_lines, accent, icon)

- 白底圆角矩形 + 顶部 accent 色条 + 0.5pt 边框
- 可选：左上角图标（Phosphor fill）+ 数字徽章
- body_lines: 文字列表，自动换行

### 箭头流程 chevron_flow(slide, steps, y, item_h, colors)

- N 个 RIGHT_ARROW 形状横排，等宽，间距 0.5%W
- 每步：主标题 + 副说明文字
- colors：可自定义每步颜色（默认循环用 ACCENTS）

### 高亮表格 htable(slide, headers, rows, x, y, w, h, hbg)

- 表头 NAVY 底白字，奇偶行 CARD/WHITE 交替
- 无边框，靠行色区分

### 标注框 callout_box(slide, x, y, w, h, text, border_color, icon)

- 左侧 border_color 竖条 + CARD 背景 + 可选图标
- 适合"注意事项/关键结论/引言"

### 分节标签 section_tag(slide, x, y, text, bg, icon)

- 圆角胶囊形标签，左侧可带图标，用于页面左上角标注章节

### 数字徽章 num_badge(slide, x, y, n, bg, size)

- 圆形 + 白色数字，用于步骤编号

---

## 技术汇报模式特殊规范

当输入为技术文档（架构说明、系统设计、代码分析等）时：

1. **封面** 用深色（NAVY 背景），右侧信息卡片显示"章节数/机制数/版本"
2. **目录** 每章节配对应语义图标
3. **设计原则页** 用 2×2 圆角卡片，每卡带图标
4. **流程/循环页** 用 chevron_flow 横排流程图
5. **工具/模块对比** 用 htable，表头 NAVY
6. **安全/风险页** 左侧红色标注框
7. **代码示例** 深色背景（DARK）文本框，代码用 SKY 色

---

## 图标使用规范

```python
from scripts.icons import get_icon

# 用法：get_icon(name, color_hex, size_px)
# 返回 PNG 文件路径，可直接用于 slide.shapes.add_picture()
icon_path = get_icon("lightbulb", "#00329D", 64)

# 在 EMU 坐标中添加图标
slide.shapes.add_picture(icon_path, Emu(x), Emu(y), Emu(size_emu), Emu(size_emu))
```

图标名称与语义对应见 `references/icon-reference.md`。

---

## 与 GordenPPTSkill 的混合使用

可以将 GordenPPTSkill 的模板页（template slides）与 Rich Mode 的原创页混合：

```python
# 从 GordenPPTSkill 模板提取若干页作为封面/结尾
src_prs = Presentation('/tmp/GordenPPTSkill/templates/SLUG/template.pptx')
tmpl_xml = [copy.deepcopy(src_prs.slides[i]._element) for i in [0, 1, -1]]

# 生成 Rich Mode 内容页
prs = new_presentation()
for mk in content_makers:
    mk(prs)

# 追加模板页
for xml_el in tmpl_xml:
    new_sl = prs.slides.add_slide(prs.slide_layouts[6])
    # ... 复制 XML
```

详见 `references/hybrid-mode.md`。

---

## 常见错误处理

| 错误 | 原因 | 解决 |
|---|---|---|
| `RGBColor() missing args` | `RGBColor(0xFFFFFF)` 单参 | 改为 `RGBColor(0xFF, 0xFF, 0xFF)` |
| 孤儿 slide 关系 | 操作了 `_sldIdLst` | 调用 `purge_orphans(prs)` |
| cairosvg 找不到 | 非标准安装路径 | `sys.path.insert(0, '/path/to/site-packages')` |
| 图标变形 | 手写 SVG path 不准 | 使用 Phosphor Icons 官方 SVG 文件 |
| 文字超出边界 | 字号或内容太长 | 减小字号或拆分到多行 |
