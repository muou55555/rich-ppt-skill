# PPT 页面类型配方手册

> 基于 GenSpark 16个HTML deck分析，总结各类幻灯片的实现配方。

---

## Type A：封面页

**适用场景**：演示文稿第1页，标题展示

```
背景: #0f172a
装饰: 右上+左下发光圆（opacity 5%）
中心: 大标题（48pt, white, 900w）
下方: 副标题（16pt, slate-400）
右下: 来源/日期标注（10pt, slate-500）
可选: 顶部 logo 区域
```

**python-pptx 实现**：
```python
slide = make_dark_slide(prs)
# 大标题居中
add_tb(slide, 0.833, 2.0, 11.667, 1.5,
       ['Claude Code 实现原理与架构'], 40,
       bold=True, color=C['text'], align=PP_ALIGN.CENTER)
# 副标题
add_tb(slide, 0.833, 3.5, 11.667, 0.5,
       ['深度解析多智能体编排与工具生态'], 16,
       color=C['muted'], align=PP_ALIGN.CENTER)
# 装饰线（居中）
add_deco_line(slide, 6.042, 1.8, w=1.25)
```

---

## Type B：目录页（两栏）

**适用场景**：演示开始前的章节索引

```
左侧(30%): 深色面板，主标题 + 说明文字 + 数字统计
右侧(70%): 2列网格，6-8个章节卡片
章节卡片: 序号(大,accent) + 标题(white) + 描述(muted)
```

**布局**：
```
┌──────────────┬───────────────────────────┐
│ 目录         │ ①章节名称   ④章节名称     │
│              │   简短描述    简短描述    │
│ 本次汇报     │                           │
│ 包含X个章节  │ ②章节名称   ⑤章节名称     │
│              │   简短描述    简短描述    │
│              │                           │
│              │ ③章节名称   ⑥章节名称     │
│              │   简短描述    简短描述    │
└──────────────┴───────────────────────────┘
```

---

## Type C：4卡片内容页

**适用场景**：展示4个并列功能/特点/原则

```
上方Header(1.1"): 装饰线 + 主标题 + 副标题 + 章节badge
下方(2×2网格): 4张功能卡片
卡片: 图标盒(0.35") + 标题(11pt,white) + 描述(9pt,muted)
```

**最佳实践**：
- 4个卡片保持相似信息密度
- 标题 ≤ 12字
- 描述 ≤ 30字
- 图标选择与内容语义匹配

---

## Type D：6卡片内容页

**适用场景**：展示6个并列功能（3×2布局）

```
上方Header(1.0"): 同Type C
下方(3×2网格): 6张功能卡片（卡片尺寸略小）
```

**注意**：字体需减小 1-2pt（标题 10pt, 正文 8pt），保证不溢出

---

## Type E：左右分栏页

**适用场景**：一侧说明 + 一侧展示（架构图/代码/截图）

```
左侧(35-40%): 
  - 蓝色装饰线
  - 大标题
  - 要点列表（圆点 + 文字）
  - 底部badge

右侧(60-65%):
  - 架构图 / 信息图 / 截图
  - 或: 内容卡片列 (垂直排列)
```

**python-pptx实现**：
```python
lx, ly, lw, lh, rx, ry, rw, rh = layout_split(slide)

# 左侧深色面板背景
panel = add_rect(slide, lx-0.1, ly-0.1, lw+0.2, lh+0.2,
                 fill=C['panel'], border=C['border'])
round_(panel, adj=6000)

# 右侧内容卡片
for i, (title, desc) in enumerate(items):
    cy = ry + i * (card_h + gap)
    add_feature_card(slide, rx, cy, rw, card_h, ...)
```

---

## Type F：时间线页

**适用场景**：展示历史进程、里程碑、发展阶段

```
上方Header(0.9"): 标题 + 时间范围副标题
左侧(55%): 垂直时间线
  - 竖线（sky-400, 2px宽, 渐变至透明）
  - 每节点: 日期(accent,11pt) + 事件(white,10pt) + 描述(muted,8pt)
右侧(45%): 数据可视化（折线图/柱状图）或补充卡片
```

**时间线节点（python-pptx）**：
```python
# 竖线
line = add_rect(slide, lx+0.12, ly, 0.02, lh, fill=C['accent'])

# 每个节点
for i, (date, title, desc) in enumerate(milestones):
    cy = ly + i * node_spacing
    
    # 节点圆点（天蓝色实心圆）
    dot = add_oval(slide, lx+0.05, cy, 0.15, 0.15, fill=C['accent'])
    
    # 日期文字
    add_tb(slide, lx+0.25, cy-0.02, 0.9, 0.2, [date], 9,
           color=C['accent'], bold=True)
    
    # 标题
    add_tb(slide, lx+0.25, cy+0.15, tw, 0.22, [title], 10,
           color=C['text'], bold=True)
    
    # 描述
    add_tb(slide, lx+0.25, cy+0.35, tw, 0.25, [desc], 8,
           color=C['muted'])
```

---

## Type G：数据统计页

**适用场景**：展示关键指标、数字成就

```
上方Header(0.8")
中间: 3-4个大数字 stat 卡片（水平排列）
  每个stat: 数字(36pt,accent,bold) + 标签(12pt,muted)
下方: 背景说明文字 or 进度条数据
```

**Stat 卡片**：
```python
# 对每个 stat
stat_w = (12.083 - gap*(n-1)) / n
card = add_rect(slide, sx, sy, stat_w, 1.5, fill=C['card'], border=C['border'])
round_(card, adj=8000)

# 大数字（accent色）
add_tb(slide, sx+0.1, sy+0.15, stat_w-0.2, 0.7,
       [value], 30, bold=True, color=C['accent'], align=PP_ALIGN.CENTER)

# 标签（muted色）
add_tb(slide, sx+0.1, sy+0.85, stat_w-0.2, 0.4,
       [label], 11, color=C['muted'], align=PP_ALIGN.CENTER)
```

---

## Type H：比较/矩阵页

**适用场景**：多方案对比、功能矩阵

```
上方Header
下方表格:
  - 表头行: bg #334155, 白色标题
  - 数据行: 交替 #1e293b / #0f172a
  - 每列等宽 or 按内容宽度
  - 状态用 ✅ ❌ ⚠️ 等 emoji 表示
```

---

## Type I：章节封面页（T页）

**适用场景**：每个大章节的过渡页

```
全屏: 
  - 左侧: 大章节序号（超大字号, 透明/半透明 accent）
  - 右下: 章节标题（32pt, white）+ 副标题（14pt, muted）
  - 底部: 蓝色宽横线（装饰）
  - 右上: 发光装饰圆
```

---

## 通用设计 Checklist

生成每张幻灯片后检查：

```
视觉层面:
□ 背景色已设置（不是白色）
□ 有装饰圆/发光效果
□ 标题上方有装饰线（deco-line）
□ 卡片有圆角
□ 图标盒背景色与图标色匹配

内容层面:
□ 标题 ≤ 20字
□ 副标题 ≤ 40字
□ 每张卡片描述 ≤ 35字
□ 卡片数量 4-6（避免超过7）

排版层面:
□ 所有元素在 7.5" 高度内（不溢出底部）
□ 所有元素在 13.333" 宽度内（不溢出右侧）
□ 左右留边 ≥ 0.4"
□ 顶部留边 ≥ 0.3"
□ 底部留边 ≥ 0.25"

字体层面:
□ 主标题 ≥ 24pt
□ 卡片标题 ≥ 10pt
□ 正文 ≥ 8pt
□ 字体设置为 Noto Sans SC
```
