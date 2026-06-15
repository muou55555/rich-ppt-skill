# 动态布局生成方法论

> 核心思想：不用固定模板，而是让 AI 根据内容结构选择最合适的布局，
> 每张 PPT 页面都是独一无二的。

---

## 一、为什么固定模板不够

传统 PPT 生成工具的问题：
- 模板固定 → 6个要点必须放6格卡片，3个要点也放6格（有的空）
- 时间线数据只能用时间线模板，即使数据少只有3个节点
- 所有页面"长一个样子"，缺乏节奏感

GenSpark 的做法：**AI 先理解内容结构，再选择匹配的布局类型，然后生成 HTML**

---

## 二、内容结构分类 → 布局选择器

```python
from enum import Enum

class ContentType(Enum):
    COVER        = "cover"         # 封面：1个主题
    TOC          = "toc"           # 目录：章节列表
    FOUR_CARDS   = "four_cards"    # 4个并列要点
    SIX_CARDS    = "six_cards"     # 6个并列要点
    SPLIT_FOCUS  = "split_focus"   # 左侧说明 + 右侧图示
    TIMELINE     = "timeline"      # 时间轴（3-6节点）
    COMPARISON   = "comparison"    # 对比/矩阵
    DATA_STATS   = "data_stats"    # 数字统计（3-4个KPI）
    FLOW_PROCESS = "flow_process"  # 流程图（3-5步骤）
    SECTION_BREAK= "section_break" # 章节过渡页

def classify_content(content_dict):
    """
    根据内容特征自动选择布局类型
    content_dict: {
        'type': str,         # 用户指定或 AI 推断
        'points': list,      # 要点数量
        'has_timeline': bool,
        'has_numbers': bool, # 包含大数据
        'has_steps': bool,   # 有步骤顺序
        'has_comparison': bool,
    }
    """
    n = len(content_dict.get('points', []))
    
    if content_dict.get('type') == 'cover': return ContentType.COVER
    if content_dict.get('type') == 'toc': return ContentType.TOC
    if content_dict.get('has_timeline'): return ContentType.TIMELINE
    if content_dict.get('has_steps') and n <= 5: return ContentType.FLOW_PROCESS
    if content_dict.get('has_comparison'): return ContentType.COMPARISON
    if content_dict.get('has_numbers') and n <= 4: return ContentType.DATA_STATS
    if n <= 4: return ContentType.FOUR_CARDS
    if n <= 6: return ContentType.SIX_CARDS
    return ContentType.SPLIT_FOCUS  # 默认：左右分栏
```

---

## 三、每种布局的 HTML 生成 Prompt

### 3.1 四卡片布局

```python
FOUR_CARDS_PROMPT = """
生成1280×720 HTML幻灯片，4卡片布局：

页面结构：
- id="slide-container" width:1280px height:720px overflow:hidden
- 背景 #0f172a
- 顶部 header 区（高100px，左内边距60px）：
  * 标题上方：蓝色装饰线 60px×4px #38bdf8
  * h1 36px 900w white：{title}
  * p 16px #94a3b8：{subtitle}
- main-content（剩余高度620px，内边距 0 60px 40px，flex row gap:30px）：
  * 2列×2行 卡片网格 gap:20px

每张卡片 CSS：
background:#1e293b; border:1px solid #334155; border-radius:12px;
padding:20px; display:flex; align-items:flex-start; gap:16px

图标盒（36×36px）：
background:rgba(56,189,248,0.1); border-radius:8px; 
display:flex; align-items:center; justify-content:center; color:#38bdf8

4个要点：
{points}

CDN（必须包含）：
- tailwindcss.com
- fontawesome 6.5.0 all.min.css
- Noto Sans SC Google Fonts
"""
```

### 3.2 时间线布局

```python
TIMELINE_PROMPT = """
生成1280×720 HTML幻灯片，左右分栏时间线布局：

左侧(55%宽)：垂直时间轴
- 左边距4px #38bdf8竖线，从上到下
- 每个节点：圆点（14px实心蓝）+ 日期（#38bdf8 accent）+ 事件标题 + 简述
- 节点间距均匀分布

右侧(45%宽)：支撑数据
- 大数字 stat（如增长率/数量）
- 或折线趋势图（用canvas绘制）

时间线数据：
{timeline_data}

装饰：
- 右上/左下半透明圆形（absolute, 300px, opacity:0.06）
"""
```

### 3.3 流程图布局

```python
FLOW_PROMPT = """
生成1280×720 HTML幻灯片，水平流程图布局：

结构：
- header（上方80px）：标题+副标题+章节badge
- 流程区（剩余空间）：{n}个步骤横向排列

每个步骤 card（宽度均分，最大220px）：
- 顶部：圆形序号（40px, bg:#38bdf8, white字, 900w）
- 箭头连接（→ 绝对定位在card间）
- 标题（16px white 700w）
- 描述（13px #94a3b8）

流程步骤：
{steps}

视觉重点：第{highlight}步用 border-color:#38bdf8 + 轻微box-shadow突出
"""
```

---

## 四、图标选择 AI Prompt

每个内容点自动匹配最合适的 FontAwesome 图标：

```python
ICON_SELECTOR_PROMPT = """
为以下技术PPT内容选择最合适的 FontAwesome 6 Solid 图标名称。
图标必须是 fa-solid 分类下的有效图标名。

内容主题：{content}

从以下候选中选择最合适的1个，只输出图标名（不含fa-）：
robot, code-branch, shield-halved, bolt, users, rocket, star, fire,
arrow-trend-up, sitemap, list-ul, globe, gem, building-columns,
chart-line, server, database, cloud, lock, check-circle,
network-wired, microchip, layer-group, puzzle-piece, wrench,
magnifying-glass, brain, infinity, gear, key, terminal,
book-open, graduation-cap, flag, medal, trophy, bullseye,
chart-bar, table-cells, filter, shuffle, copy, sync-alt
"""

def select_icon(content_text, ai_client):
    """调用 AI 为内容选择图标"""
    response = ai_client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=20,
        messages=[{
            "role": "user",
            "content": ICON_SELECTOR_PROMPT.format(content=content_text)
        }]
    )
    return response.content[0].text.strip()
```

---

## 五、完整 AI 驱动 PPT 生成流程

```python
import anthropic
from pptx import Presentation

client = anthropic.Anthropic()

def generate_rich_pptx(outline):
    """
    outline: [{
        'type': 'cover'|'content'|'section'|...,
        'title': str,
        'subtitle': str,
        'points': [{'title':str, 'desc':str}, ...],
        'notes': str,  # 演讲备注
    }]
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    for slide_plan in outline:
        # 1. 分类内容
        content_type = classify_content(slide_plan)
        
        # 2. 为每个要点选择图标
        for point in slide_plan.get('points', []):
            point['icon'] = select_icon(point['title'], client)
        
        # 3. AI 生成 HTML
        prompt = build_html_prompt(slide_plan, content_type)
        html_response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}]
        )
        html_content = extract_html(html_response.content[0].text)
        
        # 4. HTML → PPTX（混合模式）
        slide = html_to_editable_pptx(html_content, prs)
        
        # 5. 添加演讲备注
        if slide_plan.get('notes'):
            slide.notes_slide.notes_text_frame.text = slide_plan['notes']
    
    return prs


def build_html_prompt(slide_plan, content_type):
    """根据内容类型选择合适的 HTML prompt 模板"""
    templates = {
        ContentType.FOUR_CARDS:   FOUR_CARDS_PROMPT,
        ContentType.TIMELINE:     TIMELINE_PROMPT,
        ContentType.FLOW_PROCESS: FLOW_PROMPT,
        ContentType.COVER:        COVER_PROMPT,
        # ... 其他类型
    }
    template = templates.get(content_type, FOUR_CARDS_PROMPT)
    
    points_text = '\n'.join([
        f"- [{p.get('icon','star')}] {p['title']}：{p['desc']}"
        for p in slide_plan.get('points', [])
    ])
    
    return template.format(
        title=slide_plan['title'],
        subtitle=slide_plan.get('subtitle', ''),
        points=points_text,
        n=len(slide_plan.get('points', [])),
    )
```

---

## 六、视觉节奏控制

专业 PPT 的一个关键原则：**每隔几页需要视觉变化，避免所有页面单调重复**。

```python
LAYOUT_RHYTHM = [
    # 建议的页面布局序列节奏
    ContentType.COVER,         # 第1页：封面（全屏大字）
    ContentType.TOC,           # 第2页：目录（左右分栏）
    ContentType.SECTION_BREAK, # 第3页：第一章标题（全屏）
    ContentType.FOUR_CARDS,    # 第4页：4卡片内容
    ContentType.SPLIT_FOCUS,   # 第5页：左右分栏（破节奏）
    ContentType.DATA_STATS,    # 第6页：大数字（强对比）
    ContentType.TIMELINE,      # 第7页：时间线（纵向叙事）
    ContentType.FLOW_PROCESS,  # 第8页：流程（横向叙事）
    ContentType.SECTION_BREAK, # 第9页：第二章标题
    ContentType.SIX_CARDS,     # 第10页：6卡片（密度变化）
    ContentType.COMPARISON,    # 第11页：对比矩阵
]

def vary_colors(base_accent='#38bdf8', slide_index=0):
    """
    每隔几页轻微变化强调色，保持视觉新鲜感
    同一主题内保持一致，章节过渡时轻微调整
    """
    accents = [
        '#38bdf8',  # sky-400（主色）
        '#60a5fa',  # blue-400（略暖）
        '#34d399',  # emerald-400（绿色强调，数据页）
        '#a78bfa',  # violet-400（紫色，技术深度页）
        '#38bdf8',  # 回到主色
    ]
    chapter = slide_index // 4
    return accents[chapter % len(accents)]
```

---

## 七、实用工具：将任意 Markdown 转为 PPT 大纲

```python
import re

def markdown_to_outline(md_text):
    """
    将 Markdown 文档自动解析为 PPT 大纲
    支持 H1(章节) / H2(页面) / bullet(要点) 结构
    """
    outline = []
    current_section = None
    current_slide = None
    
    for line in md_text.split('\n'):
        line = line.strip()
        if not line: continue
        
        if line.startswith('# '):
            # H1 → 章节封面
            if current_slide:
                outline.append(current_slide)
            current_section = line[2:]
            outline.append({
                'type': 'section_break',
                'title': current_section,
                'subtitle': '',
                'points': []
            })
            current_slide = None
            
        elif line.startswith('## '):
            # H2 → 新页面
            if current_slide:
                outline.append(current_slide)
            current_slide = {
                'type': 'content',
                'title': line[3:],
                'subtitle': current_section or '',
                'points': []
            }
            
        elif line.startswith('- ') or line.startswith('* '):
            # bullet → 要点
            text = line[2:]
            if ':' in text:
                title, desc = text.split(':', 1)
            else:
                title, desc = text, ''
            if current_slide:
                current_slide['points'].append({
                    'title': title.strip(),
                    'desc': desc.strip()
                })
    
    if current_slide:
        outline.append(current_slide)
    
    return outline

# 使用
md = """
# Claude Code 实现原理

## 多智能体架构
- 任务分解：将复杂任务拆分为可并行执行的子任务
- 上下文共享：所有agent共享同一工具调用历史
- 结果聚合：智能合并多个agent的输出

## 工具生态
- Bash工具：直接执行系统命令，支持管道和重定向
- 文件工具：Read/Write/Edit 三件套，原子操作
- 网络工具：WebFetch/WebSearch 获取实时信息
"""

outline = markdown_to_outline(md)
prs = generate_rich_pptx(outline)
prs.save('claude_code_arch.pptx')
```

---

## 八、品质检验：自动验证生成的 PPTX

```python
def validate_pptx(pptx_path):
    """
    自动验证 PPTX 文件质量
    返回 issues 列表
    """
    from pptx import Presentation
    from pptx.util import Inches
    
    prs = Presentation(pptx_path)
    issues = []
    
    SLIDE_W = float(prs.slide_width) / 914400  # EMU → inches
    SLIDE_H = float(prs.slide_height) / 914400
    
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            left = float(shape.left) / 914400
            top = float(shape.top) / 914400
            right = left + float(shape.width) / 914400
            bottom = top + float(shape.height) / 914400
            
            # 检查溢出
            if right > SLIDE_W + 0.01:
                issues.append(f"Slide {i+1}: '{shape.name}' overflows right ({right:.3f}\" > {SLIDE_W}\")")
            if bottom > SLIDE_H + 0.01:
                issues.append(f"Slide {i+1}: '{shape.name}' overflows bottom ({bottom:.3f}\" > {SLIDE_H}\")")
            
            # 检查文字太小
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.size and run.font.size < Pt(7):
                            issues.append(f"Slide {i+1}: font too small ({run.font.size.pt:.1f}pt): '{run.text[:20]}'")
        
        # 检查是否有 notes
        notes = slide.notes_slide.notes_text_frame.text
        if not notes:
            issues.append(f"Slide {i+1}: missing speaker notes")
    
    return issues

# 验证
issues = validate_pptx('output.pptx')
if issues:
    print("⚠️ 发现问题：")
    for issue in issues:
        print(f"  - {issue}")
else:
    print("✅ 验证通过")
```
