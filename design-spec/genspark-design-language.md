# GenSpark HTML Deck 设计语言规范

> 来源：对 GenSpark AI PPT 生成产品的 16 个 HTML deck 源码进行深度分析（2024年）
> 分析页面：https://www.genspark.ai/agents?id=4300cb30-e944-4e23-ab2b-1100b4322dd9

---

## 一、核心设计原则

1. **暗色优先 (Dark-First)**：所有 deck 统一使用深蓝暗色系，背景 `#0f172a`，与技术/科技感完全一致
2. **层次清晰 (Clear Hierarchy)**：三层视觉权重——标题 → 卡片 → 辅助信息，每层颜色/尺寸明确区分
3. **蓝色强调 (Sky Blue Accent)**：单一强调色 `#38bdf8`，用于标题装饰线、图标、悬停边框、高亮文本
4. **密度适中 (Balanced Density)**：每页 4-7 个信息单元，避免超载；用留白区分区块
5. **视觉装饰轻量化**：仅用发光圆（bg-glow）和角落圆形（decor-circle）作装饰，不过度

---

## 二、配色系统

### 核心调色板

| 角色 | Hex | Tailwind 类 | 用途 |
|------|-----|-------------|------|
| 主背景 | `#0f172a` | `slate-900` | 整页背景、左侧面板 |
| 卡片背景 | `#1e293b` | `slate-800` | 内容卡片、面板背景 |
| 边框/分隔 | `#334155` | `slate-700` | 卡片边框、分隔线 |
| 次级背景 | `#475569` | `slate-600` | 悬浮状态、次级面板 |
| 主要文本 | `#f1f5f9` | `slate-100` | 标题、强调文本 |
| 次要文本 | `#cbd5e1` | `slate-300` | 副标题、说明 |
| 辅助文本 | `#94a3b8` | `slate-400` | 正文、描述 |
| 弱化文本 | `#64748b` | `slate-500` | 元数据、时间戳 |
| 强调蓝 | `#38bdf8` | `sky-400` | 装饰线、图标、高亮 |
| 强调蓝深 | `#0ea5e9` | `sky-500` | 发光效果色源 |
| 链接/按钮 | `#3b82f6` | `blue-500` | 行动元素 |
| 危险/警告 | `#ef4444` | `red-500` | 错误、风险标注 |
| 成功 | `#22c55e` | `green-500` | 正向指标 |

### 图标色背景（Icon Box）

```css
background-color: rgba(56, 189, 248, 0.10);  /* sky-400 @ 10% */
color: #38bdf8;
```

### 卡片悬停状态

```css
border-color: #38bdf8;
background-color: rgba(56, 189, 248, 0.05);
transform: translateX(-5px);
```

---

## 三、字体系统

### 字体族

```css
font-family: 'Noto Sans SC', 'PingFang SC', sans-serif;
```
> 中文首选 Noto Sans SC，英文天然 fallback 至 system-sans

### 字体大小等级（16个deck统计）

| 级别 | px 值 | 用途 | 权重 |
|------|-------|------|------|
| Display | 48-60px | 封面主标题 | 900 |
| H1 | 36px | 页面标题 | 900 |
| H2 | 28-30px | 左侧面板主标题 | 800 |
| H3 | 20-22px | 卡片分组标题 | 700 |
| Body-L | 16px | 副标题、说明文字 | 400 |
| Body | 15px | 卡片标题 | 700 |
| Body-S | 14px | 卡片内容 | 400 |
| Caption | 12px | 标签、注释 | 400 |
| Mono | 13px | 代码、技术标注 | 400 (monospace) |

### Letter Spacing

```css
/* 大标题防粘字 */
h1 { letter-spacing: -0.02em; }
/* 小标签增强可读性 */
.badge { letter-spacing: 0.05em; text-transform: uppercase; }
```

---

## 四、页面布局系统

### 画布尺寸

```css
width: 1280px;
height: 720px;
overflow: hidden;
position: relative;
```

> python-pptx 对应：宽 13.333 英寸，高 7.5 英寸（16:9）

### 通用两区分割

```
┌─────────────────────────────────────────┐ 720px
│  [装饰圆-右上]     [装饰圆-左下]            │
│  ┌──────────────┐  ┌──────────────────┐  │
│  │  Header      │  Header cont.        │  │
│  │  30px top    │  100px height        │  │
│  └──────────────┘  └──────────────────┘  │
│  ┌──────────────┐  ┌──────────────────┐  │
│  │  left-panel  │  │   right-panel    │  │
│  │  ~380px      │  │   flex: 1        │  │
│  │              │  │  (卡片网格 or     │  │
│  │              │  │   时间线等)       │  │
│  └──────────────┘  └──────────────────┘  │
└─────────────────────────────────────────┘ 1280px
```

### Header 区块规范

```css
.header {
  padding: 30px 60px 10px;
  height: 100px;
  display: flex;
  align-items: flex-end;
  justify-content: space-between;
  flex-shrink: 0;
}
```

### Main Content 区块规范

```css
.main-content {
  flex: 1;
  display: flex;
  padding: 0 60px 40px;
  gap: 30px;
  height: calc(100% - 100px);
  position: relative;
  z-index: 10;
}
```

### Left Panel 规范

```css
.left-panel {
  width: 380px;
  flex-shrink: 0;
  background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
  border-radius: 16px;
  padding: 32px 28px;
  border: 1px solid #334155;
  position: relative;
  overflow: hidden;
}
```

---

## 五、核心组件规范

### 5.1 装饰线（Deco Line）

标题上方的蓝色短横线，是 GenSpark 最标志性的视觉元素之一。

```css
.deco-line {
  width: 60px;
  height: 4px;
  background-color: #38bdf8;
  border-radius: 2px;
  margin-bottom: 8px;
}
```

**python-pptx 实现：**
```python
line = slide.shapes.add_shape(1, E(1.0), E(0.4), E(0.625), E(0.042))
line.fill.solid(); line.fill.fore_color.rgb = RGBColor(0x38, 0xBD, 0xF8)
line.line.fill.background()
```

### 5.2 发光背景（BG Glow）

两个大圆形，高斯模糊，制造科技感光晕。

```css
.bg-glow {
  position: absolute;
  width: 600px;
  height: 600px;
  background-color: #0ea5e9;
  filter: blur(120px);
  opacity: 0.05;
  border-radius: 50%;
  z-index: 0;
  pointer-events: none;
}
/* 一个置于右上，一个置于左下 */
.bg-glow:first-child { top: -200px; right: -200px; }
.bg-glow:last-child { bottom: -200px; left: -200px; }
```

**python-pptx 实现（近似）：**
> python-pptx 不支持 blur，用低透明度大圆形替代：
```python
oval = slide.shapes.add_shape(9, E(-1.5), E(-1.5), E(6.25), E(6.25))
oval.fill.solid()
oval.fill.fore_color.rgb = RGBColor(0x0E, 0xA5, 0xE9)
oval.fill.fore_color.theme_color = None
# 设置透明度
from pptx.util import Pt
from lxml import etree
# 直接用 XML 设置透明度
sp = oval._element
solidFill = sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill')
# 透明度 90% = lumMod 10000
```

### 5.3 角落装饰圆（Decor Circle）

右上角和左下角的半透明圆，增加页面深度感。

```css
.decor-circle {
  position: absolute;
  width: 300px;
  height: 300px;
  border-radius: 50%;
  background: radial-gradient(circle, rgba(56,189,248,0.08) 0%, transparent 70%);
}
.decor-circle:nth-child(1) { top: -80px; right: -80px; }
.decor-circle:nth-child(2) { bottom: -80px; left: -80px; }
```

### 5.4 功能卡片（Feature Card）

最常用的内容单元，4-6 张组成网格。

```css
.feature-card {
  background-color: #1e293b;
  border: 1px solid #334155;
  border-radius: 12px;
  padding: 12px 16px;
  display: flex;
  align-items: flex-start;
  gap: 12px;
  transition: all 0.3s ease;
}
.feature-card:hover {
  transform: translateX(-5px);
  border-color: #38bdf8;
  background-color: rgba(56, 189, 248, 0.05);
}
```

**卡片内部结构：**
```
[icon-box 36×36]  [content]
                  [h3 15px 700 white]
                  [p 12px slate-400 1.4lh]
```

### 5.5 图标盒（Icon Box）

配合 FontAwesome 使用，蓝色透明背景。

```css
.icon-box {
  width: 36px;
  height: 36px;
  border-radius: 8px;
  background-color: rgba(56, 189, 248, 0.10);
  display: flex;
  align-items: center;
  justify-content: center;
  color: #38bdf8;
  font-size: 16px;
  flex-shrink: 0;
}
```

**python-pptx 实现：**
```python
# 小圆角矩形 + emoji 或 Unicode 图标文字
icon_box = slide.shapes.add_shape(1, ix, iy, E(0.375), E(0.375))
icon_box.fill.solid()
icon_box.fill.fore_color.rgb = RGBColor(0x38, 0xBD, 0xF8)
icon_box.fill.fore_color.theme_color = None
# 设置 15% 透明度 → 用 XML lumMod
round_(icon_box, adj=8000)  # 圆角 ~8% radius
# 文字用 emoji: 🤖 📋 🔒 ⚡ 🔧 📊 🌐 等
```

### 5.6 标签/徽章（Badge / Pill）

小型圆角矩形，用于分类标注。

```css
.badge {
  display: inline-flex;
  align-items: center;
  padding: 4px 12px;
  border-radius: 9999px;  /* pill */
  background-color: rgba(56, 189, 248, 0.15);
  color: #38bdf8;
  font-size: 12px;
  font-weight: 600;
  letter-spacing: 0.05em;
  border: 1px solid rgba(56, 189, 248, 0.3);
}
```

**变体：**
- `slate` badge: bg `#334155`, text `#94a3b8`
- `sky` badge: bg rgba(sky, 0.15), text `#38bdf8`
- `red` badge: bg rgba(red, 0.15), text `#ef4444`
- `green` badge: bg rgba(green, 0.15), text `#22c55e`

**python-pptx 实现：**
```python
pill = slide.shapes.add_shape(1, px, py, pw, E(0.25))
pill.fill.solid(); pill.fill.fore_color.rgb = RGBColor(0x1E, 0x29, 0x3B)
round_(pill, adj=50000)  # 完全圆角 = pill 形状
```

### 5.7 时间线（Timeline）

垂直时间轴，每个里程碑一行。

```css
.timeline-container {
  display: flex;
  flex-direction: column;
  gap: 16px;
  position: relative;
  padding-left: 24px;
}
/* 左侧竖线 */
.timeline-container::before {
  content: '';
  position: absolute;
  left: 6px;
  top: 0; bottom: 0;
  width: 2px;
  background: linear-gradient(to bottom, #38bdf8, transparent);
}
/* 圆点 */
.timeline-dot {
  width: 14px; height: 14px;
  border-radius: 50%;
  background: #38bdf8;
  border: 2px solid #0f172a;
  position: absolute;
  left: -5px;
}
```

### 5.8 进度条（Progress Bar）

横向条形，用于展示指标比例。

```css
.progress-track {
  height: 6px;
  background: #334155;
  border-radius: 3px;
}
.progress-fill {
  height: 100%;
  border-radius: 3px;
  background: linear-gradient(90deg, #0ea5e9, #38bdf8);
}
/* 宽度用内联 style: width: 75% */
```

### 5.9 节点徽章（Layer Badge）

左侧面板顶部的章节/类别标注。

```css
.layer-badge {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 6px 14px;
  background: rgba(14, 165, 233, 0.15);
  border: 1px solid rgba(14, 165, 233, 0.4);
  border-radius: 20px;
  font-size: 12px;
  color: #38bdf8;
  font-weight: 600;
  margin-bottom: 20px;
}
```

---

## 六、页面类型模板

### Type A：封面页（Cover Slide）

```
背景: #0f172a
左上: 发光圆装饰
中央: 大标题 (48-60px, white, 900w)
下方: 副标题 (20px, slate-400)
右下: 角落圆装饰
```

### Type B：目录页（TOC Slide）

```
左 30%: left-panel (深色，TOC 编号列表)
右 70%: right-panel (grid-cols-2, 卡片网格)
每章节: 序号 + 标题 + 简短描述
```

### Type C：内容卡片页（Card Grid Slide）

```
上方: header (标题 + 副标题 + 装饰线 + 徽章)
下方: main-content (2-3 列卡片网格)
卡片: icon-box + 标题 + 描述
```

### Type D：分析页（Split Analysis Slide）

```
左半: 文字列表/说明/关键点
右半: 架构图/流程图/数据可视化 (canvas 或 SVG)
分隔: 1px solid #334155
```

### Type E：时间线/流程页（Timeline Slide）

```
上方: header
下方左: timeline-container (垂直时间轴)
下方右: chart-wrapper (趋势图/增长图)
```

### Type F：矩阵/比较页（Matrix Slide）

```
表头: 颜色行 (#334155 bg)
内容行: 交替 #1e293b / #0f172a
单元格: 居中文字 + 彩色状态点
```

---

## 七、图标使用规范

所有图标使用 **FontAwesome 6 Solid**（`fa-solid` + `fa-xxx`）。

### 常用图标映射

| 语义 | FA 图标 | Unicode |
|------|---------|---------|
| AI/机器人 | `fa-robot` | 🤖 |
| 代码/分支 | `fa-code-branch` | 🔀 |
| 安全/盾牌 | `fa-shield-halved` | 🛡️ |
| 闪电/速度 | `fa-bolt` | ⚡ |
| 用户群 | `fa-users` | 👥 |
| 趋势 | `fa-arrow-trend-up` | 📈 |
| 火/热点 | `fa-fire` | 🔥 |
| 星/评级 | `fa-star` | ⭐ |
| 火箭 | `fa-rocket` | 🚀 |
| 网格/模块 | `fa-sitemap` | 🗂️ |
| 文件 | `fa-list-ul` | 📋 |
| 历史 | `fa-history` | 🕐 |
| 全球/网络 | `fa-globe` | 🌐 |
| 宝石/高质量 | `fa-gem` | 💎 |
| 建筑/机构 | `fa-building-columns` | 🏛️ |

> python-pptx 中用 Emoji 代替 FA 图标（Segoe UI Emoji 字体）

---

## 八、HTML → python-pptx 映射表

| HTML 元素/概念 | python-pptx 实现 |
|----------------|-----------------|
| `<div class="slide-container">` | Slide (13.333" × 7.5") |
| `background-color: #0f172a` | `slide.background.fill.solid()` |
| `<div class="header">` | TextBox at top, h=1.042" (100px/96dpi) |
| `<h1 style="font-size:36px">` | TextBox, font size 27pt (~36px) |
| `.deco-line` 60×4px | `add_shape(1, ...)`, H=0.042", W=0.625" |
| `<div class="feature-card">` | `add_shape(1,...)` + `round_(adj=8000)` |
| `.icon-box 36×36px` | `add_shape(1,...)` W/H=0.375", `round_(adj=8000)` |
| `.badge pill` | `add_shape(1,...)` + `round_(adj=50000)` |
| `.bg-glow 600px blur` | 椭圆 `add_shape(9,...)`, 低透明度, 无线条 |
| `.decor-circle` | 椭圆 `add_shape(9,...)`, 低透明度渐变 |
| `font-size: 36px` → pt | px × (72/96) = pt，即 36px = 27pt |
| `1280px → 13.333"` | 1px = 13.333/1280 = 0.01042" |
| `720px → 7.5"` | 1px = 7.5/720 = 0.01042" |

### EMU 换算速查

```python
def px_to_emu(px):
    return int(px * 914400 / 96)  # 96dpi screen

def pt_to_emu(pt):
    return int(pt * 12700)

# 常用值
# 60px deco-line width  = 572,625 EMU
# 4px deco-line height  = 38,175 EMU
# 36px = 34pt font size
# 100px header height   = 952,500 EMU
# 60px side padding     = 571,500 EMU
```

---

## 九、python-pptx 关键工具函数

以下函数是将 GenSpark 设计转为 PPTX 的核心工具集：

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from lxml import etree

# 命名空间辅助
def qn(tag):
    nsmap = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }
    prefix, local = tag.split(':')
    return f'{{{nsmap[prefix]}}}{local}'

# 英寸转 EMU（主要单位工具）
def E(inch): return int(inch * 914400)

# 像素转英寸（1280px = 13.333"）
def px(p): return p * 13.333 / 1280

# 添加矩形（fill 和 border 可选）
def add_rect(slide, x, y, w, h, fill=None, border=None, bpt=0.75):
    s = slide.shapes.add_shape(1, E(x), E(y), E(w), E(h))
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if border:
        s.line.color.rgb = border
        s.line.width = Pt(bpt)
    else:
        s.line.fill.background()
    return s

# 圆角化（将矩形转为 roundRect）
def round_(shape, adj=16667):
    """adj: 0=直角, 16667≈10%, 50000=pill"""
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    pg = spPr.find(qn('a:prstGeom'))
    pg.set('prst', 'roundRect')
    al = pg.find(qn('a:avLst'))
    if al is None:
        al = etree.SubElement(pg, qn('a:avLst'))
    for g in list(al):
        al.remove(g)
    gd = etree.SubElement(al, qn('a:gd'))
    gd.set('name', 'adj')
    gd.set('fmla', f'val {adj}')

# 添加文本框（支持多行、行距、颜色）
def add_tb(slide, x, y, w, h, lines, size, bold=False,
           color=None, align=PP_ALIGN.LEFT, ls=1.15):
    txBox = slide.shapes.add_textbox(E(x), E(y), E(w), E(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        # 行距
        p.line_spacing = Pt(size * ls)
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color or RGBColor(0xF1, 0xF5, 0xF9)
        run.font.name = 'Noto Sans SC'
    return txBox

# 设置幻灯片背景色
def set_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

# 常用颜色常量
C = {
    'bg':         RGBColor(0x0F, 0x17, 0x2A),  # slate-900
    'card':       RGBColor(0x1E, 0x29, 0x3B),  # slate-800
    'border':     RGBColor(0x33, 0x41, 0x55),  # slate-700
    'accent':     RGBColor(0x38, 0xBD, 0xF8),  # sky-400
    'accent2':    RGBColor(0x0E, 0xA5, 0xE9),  # sky-500
    'text':       RGBColor(0xF1, 0xF5, 0xF9),  # slate-100
    'text2':      RGBColor(0xCB, 0xD5, 0xE1),  # slate-300
    'muted':      RGBColor(0x94, 0xA3, 0xB8),  # slate-400
    'faint':      RGBColor(0x64, 0x74, 0x8B),  # slate-500
    'danger':     RGBColor(0xEF, 0x44, 0x44),  # red-500
    'success':    RGBColor(0x22, 0xC5, 0x5E),  # green-500
    'icon_bg':    RGBColor(0x0A, 0x1A, 0x2F),  # icon box bg (near slate-900)
}
```

---

## 十、设计质量检查清单

在生成每张幻灯片时，检查以下项目：

- [ ] 背景色是否为 `#0f172a`（或指定主题色）
- [ ] 标题上方是否有装饰线（deco-line）
- [ ] 标题字号 ≥ 27pt（对应 36px）
- [ ] 正文字号 ≥ 10pt（对应 14px）
- [ ] 所有内容是否在 7.5" 高度内（无溢出）
- [ ] 卡片是否有圆角（adj ≥ 5000）
- [ ] 图标盒是否有轻微圆角（adj ≥ 8000）
- [ ] Badge/Pill 是否完全圆角（adj = 50000）
- [ ] 是否有角落装饰圆（增加层次感）
- [ ] 颜色对比度：文字 vs 背景 ≥ 4.5:1
- [ ] 关键数据是否用强调色（`#38bdf8`）突出
- [ ] 卡片数量是否适中（4-7 个，避免密度过高）

---

## 十一、页面密度规范

| 内容量 | 推荐布局 | 单元数 |
|--------|----------|--------|
| 1 个主题 | 封面/过渡页 | 1 |
| 2-3 个要点 | 左右分栏 | 2-3 |
| 4 个功能 | 2×2 卡片网格 | 4 |
| 6 个功能 | 2×3 或 3×2 卡片网格 | 6 |
| 时间线 | 垂直列 + 数据图 | 3-5 节点 |
| 对比矩阵 | 表格 | ≤ 4 列 × 5 行 |

---

## 十二、动画/过渡规范（HTML 端）

GenSpark HTML deck 使用的动画：

```css
/* 卡片悬停滑动 */
transition: all 0.3s ease;
transform: translateX(-5px);

/* 数字闪烁 */
@keyframes countUp { from { opacity:0; } to { opacity:1; } }

/* 进度条填充 */
@keyframes fillBar { from { width:0; } to { width: var(--target); } }
```

> python-pptx 不支持 CSS 动画，PPTX 中可用 PowerPoint 原生动画（MorphTransition）替代，或省略。
