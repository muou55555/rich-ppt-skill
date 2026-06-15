# HTML Deck → 可编辑 PPTX 完整方法论

> GenSpark 的核心思路：先用 HTML + 在线库生成任意视觉效果，再通过解析/渲染管道转为可编辑 PPTX。
> 这套方法论不依赖固定模板，每张幻灯片的布局、图标、配色都可以由 AI 动态决定。

---

## 一、GenSpark HTML Deck 的在线依赖库

每个 HTML deck 的 `<head>` 区域包含以下 CDN：

```html
<!-- 1. Tailwind CSS — 零配置工具类，无需本地编译 -->
<script src="https://cdn.tailwindcss.com"></script>

<!-- 2. FontAwesome 6 Solid — 图标库，约 2000+ 图标 -->
<link rel="stylesheet"
  href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">

<!-- 3. Google Fonts — Noto Sans SC 支持中文 -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;700;900&display=swap"
  rel="stylesheet">

<!-- 4. 可选：Chart.js 数据图表 -->
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
```

**为什么这套组合能生成丰富页面：**
- **Tailwind**：500+ 工具类直接写在 class 属性上，无需自定义 CSS，AI 生成时极其方便
- **FontAwesome**：`<i class="fa-solid fa-robot"></i>` 一行代码即可嵌入精美图标
- **Noto Sans SC**：唯一能优雅渲染中英混排的免费字体
- **Canvas API**（原生）：折线图、柱状图、甜甜圈图不依赖任何库

---

## 二、FontAwesome 图标在 PPTX 中的三种实现方案

### 方案 A：下载 SVG 文件嵌入（★ 推荐）

FontAwesome SVG 文件可直接从 CDN 下载，然后作为图片嵌入 PPTX。

```python
import requests
from io import BytesIO
from pptx.util import Inches, Pt, Emu

FA_SVG_BASE = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/svgs/solid/{icon}.svg"

def get_fa_icon_png(icon_name, color_hex="#38bdf8", size_px=64):
    """
    下载 FA solid SVG，修改颜色，转为 PNG（需要 cairosvg）
    icon_name: 如 'robot', 'shield-halved', 'bolt'
    """
    import cairosvg
    url = FA_SVG_BASE.format(icon=icon_name)
    resp = requests.get(url, timeout=5)
    svg_content = resp.text
    # 替换填充色（FA SVG 默认黑色）
    svg_colored = svg_content.replace('<svg ', f'<svg fill="{color_hex}" ')
    # 转 PNG
    png_bytes = cairosvg.svg2png(bytestring=svg_colored.encode(),
                                  output_width=size_px, output_height=size_px)
    return BytesIO(png_bytes)

def add_fa_icon(slide, icon_name, x, y, size_inch=0.4, color="#38bdf8"):
    """在幻灯片中嵌入 FontAwesome 图标"""
    png_stream = get_fa_icon_png(icon_name, color_hex=color, size_px=128)
    pic = slide.shapes.add_picture(png_stream,
                                    Inches(x), Inches(y),
                                    Inches(size_inch), Inches(size_inch))
    return pic

# 使用示例
add_fa_icon(slide, 'robot', x=1.0, y=2.0, size_inch=0.35, color='#38bdf8')
add_fa_icon(slide, 'shield-halved', x=4.0, y=2.0, size_inch=0.35)
add_fa_icon(slide, 'bolt', x=7.0, y=2.0, size_inch=0.35)
```

**常用 Solid 图标名称（直接用于 URL）：**
```
robot          shield-halved    bolt           code-branch
users          rocket           star           fire
arrow-trend-up sitemap         list-ul        globe
gem            building-columns chart-line     server
database       cloud            lock          check-circle
xmark          circle-info      triangle-excl. magnifying-glass
```

### 方案 B：FontAwesome Unicode + 字体文件

```python
# 需要系统安装或项目包含 FA 字体文件
# FontAwesome Solid Unicode 码点示例
FA_UNICODE = {
    'robot':         '',
    'shield-halved': '',
    'bolt':          '',
    'code-branch':   '',
    'users':         '',
    'rocket':        '',
    'star':          '',
    'fire':          '',
    'globe':         '',
    'lock':          '',
    'database':      '',
    'server':        '',
    'cloud':         '',
    'check-circle':  '',
}

def add_fa_text_icon(slide, icon_name, x, y, size_inch=0.4, color=None):
    from pptx.dml.color import RGBColor
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(size_inch), Inches(size_inch))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = FA_UNICODE.get(icon_name, '?')
    run.font.name = 'Font Awesome 6 Free'  # 需要安装字体
    run.font.size = Pt(size_inch * 72 * 0.7)
    run.font.color.rgb = color or RGBColor(0x38, 0xBD, 0xF8)
```

### 方案 C：Emoji 替代（无依赖，快速）

```python
# AI 选择最接近语义的 emoji，无需外部依赖
ICON_EMOJI = {
    'robot':         '🤖',
    'shield-halved': '🛡️',
    'bolt':          '⚡',
    'code-branch':   '🔀',
    'users':         '👥',
    'rocket':        '🚀',
    'star':          '⭐',
    'fire':          '🔥',
    'globe':         '🌐',
    'database':      '🗄️',
    'server':        '🖥️',
    'chart-line':    '📈',
    'lock':          '🔒',
    'check-circle':  '✅',
    'sitemap':       '🗂️',
}
```

---

## 三、HTML → PPTX 的三种转换模式

### 模式 1：截图嵌入（Playwright → 图片）

**适用**：视觉最完美，但文字不可直接编辑

```python
from playwright.sync_api import sync_playwright
from pptx.util import Inches, Emu
from io import BytesIO

def html_slide_to_png(html_content, output_path=None):
    """
    用 Playwright 将 HTML slide 截图为 PNG
    返回 BytesIO 流
    """
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={'width': 1280, 'height': 720})
        page.set_content(html_content)
        # 等待字体和图标加载完成
        page.wait_for_load_state('networkidle')
        
        png_bytes = page.screenshot(
            type='png',
            clip={'x': 0, 'y': 0, 'width': 1280, 'height': 720}
        )
        browser.close()
    
    stream = BytesIO(png_bytes)
    if output_path:
        with open(output_path, 'wb') as f:
            f.write(png_bytes)
    return stream

def embed_png_as_slide(prs, png_stream):
    """将 PNG 作为全页图片嵌入 PPTX"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(png_stream, 0, 0,
                              prs.slide_width, prs.slide_height)
    return slide
```

**安装依赖：**
```bash
pip install playwright cairosvg
playwright install chromium
```

---

### 模式 2：混合模式（截图背景 + 文本叠加）★ GenSpark 实际采用

**思路**：
1. Playwright 截图 → 背景图（含所有图标、色块、装饰）
2. 解析 HTML DOM 中的文本元素 → 叠加可编辑 TextBox
3. 结果：视觉完全一致，文字可编辑

```python
from bs4 import BeautifulSoup
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

SLIDE_W_IN = 13.333  # inches
SLIDE_H_IN = 7.5     # inches

def px_to_in_x(px): return px * SLIDE_W_IN / 1280
def px_to_in_y(px): return px * SLIDE_H_IN / 720
def E(inch): return int(inch * 914400)

def get_computed_bounds(el_style):
    """从 style 字符串解析位置和尺寸（简化版）"""
    import re
    def get_val(prop):
        m = re.search(rf'{prop}:\s*([\d.]+)px', el_style)
        return float(m.group(1)) if m else None
    return {k: get_val(k) for k in ['left', 'top', 'width', 'height']}

def html_to_editable_pptx(html_content, prs):
    """
    混合模式：截图 + 文本叠加
    """
    # Step 1: 截图整张幻灯片
    png_stream = html_slide_to_png(html_content)
    slide = embed_png_as_slide(prs, png_stream)
    
    # Step 2: 用 Playwright 获取文本元素的实际渲染位置
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={'width': 1280, 'height': 720})
        page.set_content(html_content)
        page.wait_for_load_state('networkidle')
        
        # 获取所有文本节点的 bounding box
        text_elements = page.evaluate("""
            () => {
                const textTags = ['H1','H2','H3','H4','P','SPAN','LI'];
                const results = [];
                document.querySelectorAll(textTags.join(',')).forEach(el => {
                    const rect = el.getBoundingClientRect();
                    if (rect.width < 5 || rect.height < 5) return;
                    const style = window.getComputedStyle(el);
                    results.push({
                        tag: el.tagName,
                        text: el.innerText.trim(),
                        x: rect.left, y: rect.top,
                        w: rect.width, h: rect.height,
                        fontSize: parseFloat(style.fontSize),
                        fontWeight: style.fontWeight,
                        color: style.color,
                        textAlign: style.textAlign,
                    });
                });
                return results;
            }
        """)
        browser.close()
    
    # Step 3: 在 PPTX 中叠加透明 TextBox
    for el in text_elements:
        if not el['text']:
            continue
        
        x_in = px_to_in_x(el['x'])
        y_in = px_to_in_y(el['y'])
        w_in = px_to_in_x(el['w'])
        h_in = px_to_in_y(el['h'])
        
        # 解析颜色 "rgb(r, g, b)" → RGBColor
        color = parse_rgb(el['color'])
        font_pt = el['fontSize'] * 72 / 96  # px → pt
        bold = int(el.get('fontWeight', 400)) >= 700
        align = {
            'center': PP_ALIGN.CENTER,
            'right': PP_ALIGN.RIGHT,
        }.get(el['textAlign'], PP_ALIGN.LEFT)
        
        txBox = slide.shapes.add_textbox(E(x_in), E(y_in), E(w_in), E(h_in))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = el['text']
        run.font.size = Pt(font_pt)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = 'Noto Sans SC'
        # 文本框本身无填充（透明）
        txBox.fill.background()
        # 无边框
        txBox.line.fill.background()
    
    return slide

def parse_rgb(css_rgb):
    """'rgb(241, 245, 249)' → RGBColor"""
    import re
    m = re.search(r'rgb\((\d+),\s*(\d+),\s*(\d+)\)', css_rgb)
    if m:
        return RGBColor(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return RGBColor(0xFF, 0xFF, 0xFF)
```

---

### 模式 3：纯 DOM 解析 → python-pptx 形状（完全可编辑）

**思路**：不截图，完全通过解析 HTML/CSS 结构来生成 PPTX 形状。
**难点**：需要实现一个简化的 CSS 布局引擎（Flexbox/Grid → 坐标）

```python
from bs4 import BeautifulSoup
import re

class HTMLSlideToPPTX:
    """
    HTML Deck → 完全可编辑 PPTX 转换器
    
    不依赖截图，通过解析 HTML 结构生成 PPTX shapes。
    支持的 CSS 属性：position, left, top, width, height,
    background-color, color, font-size, font-weight, 
    border-radius, opacity, border-color
    """
    
    TAILWIND_COLORS = {
        'slate-900': '#0f172a', 'slate-800': '#1e293b', 'slate-700': '#334155',
        'slate-600': '#475569', 'slate-500': '#64748b', 'slate-400': '#94a3b8',
        'slate-300': '#cbd5e1', 'slate-200': '#e2e8f0', 'slate-100': '#f1f5f9',
        'sky-400':   '#38bdf8', 'sky-500':   '#0ea5e9', 'sky-600':   '#0284c7',
        'blue-500':  '#3b82f6', 'blue-400':  '#60a5fa', 'blue-300':  '#93c5fd',
        'red-500':   '#ef4444', 'green-500': '#22c55e', 'amber-500': '#f59e0b',
        'violet-500':'#a855f7', 'white':     '#ffffff', 'black':     '#000000',
    }
    
    def __init__(self, prs):
        self.prs = prs
        self.slide = None
    
    def convert(self, html_content):
        """主入口：HTML string → PPTX slide"""
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # 创建幻灯片
        slide_layout = self.prs.slide_layouts[6]
        self.slide = self.prs.slides.add_slide(slide_layout)
        
        # 解析 <style> 标签中的 CSS 规则
        style_text = '\n'.join(s.get_text() for s in soup.find_all('style'))
        self.css_rules = self._parse_css(style_text)
        
        # 从 slide-container 开始递归处理
        container = soup.find(id='slide-container') or soup.find('body')
        self._process_element(container, offset_x=0, offset_y=0)
        
        return self.slide
    
    def _process_element(self, el, offset_x, offset_y):
        """递归处理元素"""
        if el.name is None:  # 文本节点
            return
        
        # 获取该元素的完整 style（标签内 + CSS 规则）
        computed = self._compute_style(el)
        bounds = self._get_bounds(computed, offset_x, offset_y)
        
        if not bounds:
            # 没有明确位置信息，处理流式布局子元素
            for child in el.children:
                self._process_element(child, offset_x, offset_y)
            return
        
        x, y, w, h = bounds
        
        # 背景矩形
        bg = computed.get('background-color')
        border = computed.get('border-color') or computed.get('border')
        if bg and bg != 'transparent' and bg != 'rgba(0, 0, 0, 0)':
            rect = self._add_bg_shape(x, y, w, h, bg, border,
                                       computed.get('border-radius', '0px'))
        
        # 文本内容
        text = el.get_text(strip=True)
        if text and el.name in ['h1','h2','h3','h4','h5','p','span','li','a']:
            self._add_text_shape(x, y, w, h, text, computed)
        
        # 图标（FA icon）
        if el.name == 'i' and el.get('class'):
            fa_classes = [c for c in el['class'] if c.startswith('fa-') and c != 'fa-solid']
            if fa_classes:
                self._add_icon(x, y, min(w,h), fa_classes[0].replace('fa-', ''))
        
        # 递归子元素
        for child in el.children:
            self._process_element(child, x, y)
    
    def _add_bg_shape(self, x, y, w, h, bg_color, border_color, border_radius):
        from pptx.dml.color import RGBColor
        s = self.slide.shapes.add_shape(1, E(x), E(y), E(w), E(h))
        rgb = hex_to_rgb(bg_color)
        if rgb:
            s.fill.solid()
            s.fill.fore_color.rgb = RGBColor(*rgb)
        else:
            s.fill.background()
        
        if border_color:
            rgb_b = hex_to_rgb(border_color)
            if rgb_b:
                s.line.color.rgb = RGBColor(*rgb_b)
                s.line.width = Pt(0.75)
            else:
                s.line.fill.background()
        else:
            s.line.fill.background()
        
        # 圆角
        radius_px = float(re.sub(r'[^\d.]', '', border_radius or '0') or 0)
        if radius_px > 0:
            # 将 px 圆角转为 adj 值（近似）
            adj = int(min(radius_px / min(w,h) * 100000, 50000))
            round_(s, adj)
        
        return s
    
    def _get_bounds(self, computed, offset_x, offset_y):
        """从 computed style 获取绝对坐标（英寸）"""
        position = computed.get('position', 'static')
        
        def css_to_in(val, axis_size):
            if not val or val == 'auto': return None
            if val.endswith('px'):
                px = float(val[:-2])
                if axis_size == 'x': return px_to_in_x(px)
                else: return px_to_in_y(px)
            if val.endswith('%'):
                pct = float(val[:-1]) / 100
                if axis_size == 'x': return pct * SLIDE_W_IN
                else: return pct * SLIDE_H_IN
            return None
        
        if position in ('absolute', 'fixed'):
            left = css_to_in(computed.get('left'), 'x')
            top  = css_to_in(computed.get('top'), 'y')
            w    = css_to_in(computed.get('width'), 'x')
            h    = css_to_in(computed.get('height'), 'y')
            if left is not None and top is not None and w and h:
                return left + offset_x, top + offset_y, w, h
        return None
    
    def _parse_css(self, style_text):
        """简单 CSS 解析器：提取 .class { prop: val } 规则"""
        rules = {}
        for match in re.finditer(r'([.#\w][^\{]*)\{([^\}]+)\}', style_text):
            selectors = match.group(1).strip().split(',')
            props = {}
            for decl in match.group(2).split(';'):
                if ':' in decl:
                    k, v = decl.split(':', 1)
                    props[k.strip()] = v.strip()
            for sel in selectors:
                rules[sel.strip()] = props
        return rules
    
    def _compute_style(self, el):
        """合并 CSS 规则 + 内联 style"""
        computed = {}
        # CSS 规则（按选择器查找）
        for cls in (el.get('class') or []):
            rule = self.css_rules.get(f'.{cls}', {})
            computed.update(rule)
        # 内联 style 优先级最高
        inline = el.get('style', '')
        for decl in inline.split(';'):
            if ':' in decl:
                k, v = decl.split(':', 1)
                computed[k.strip()] = v.strip()
        return computed
    
    def _add_icon(self, x, y, size_in, icon_name):
        """嵌入 FA 图标（方案A：下载SVG）"""
        try:
            png = get_fa_icon_png(icon_name, color_hex='#38bdf8', size_px=128)
            self.slide.shapes.add_picture(png, E(x), E(y), E(size_in), E(size_in))
        except:
            # 降级到 emoji
            emoji = ICON_EMOJI.get(icon_name, '●')
            txBox = self.slide.shapes.add_textbox(E(x), E(y), E(size_in), E(size_in))
            tf = txBox.text_frame
            run = tf.paragraphs[0].add_run()
            run.text = emoji
            run.font.size = Pt(size_in * 72 * 0.8)
    
    def _add_text_shape(self, x, y, w, h, text, computed):
        from pptx.dml.color import RGBColor
        size_px = float(re.sub(r'[^\d.]', '', computed.get('font-size','14px')) or 14)
        size_pt = size_px * 72 / 96
        weight = computed.get('font-weight', '400')
        bold = int(weight) >= 700 if weight.isdigit() else weight in ('bold','bolder')
        color_hex = computed.get('color', '#f1f5f9')
        rgb = hex_to_rgb(color_hex) or (241,245,249)
        
        txBox = self.slide.shapes.add_textbox(E(x), E(y), E(w), E(h))
        tf = txBox.text_frame
        tf.word_wrap = True
        txBox.fill.background()
        txBox.line.fill.background()
        run = tf.paragraphs[0].add_run()
        run.text = text
        run.font.size = Pt(size_pt)
        run.font.bold = bold
        run.font.color.rgb = RGBColor(*rgb)
        run.font.name = 'Noto Sans SC'


# ─── 辅助函数 ──────────────────────────────────────────────────────────────

def hex_to_rgb(color_str):
    """'#38bdf8' or 'rgb(56,189,248)' → (56,189,248)"""
    if not color_str: return None
    color_str = color_str.strip()
    m = re.match(r'#([0-9a-fA-F]{6})', color_str)
    if m:
        h = m.group(1)
        return int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
    m = re.match(r'rgb\((\d+),\s*(\d+),\s*(\d+)\)', color_str)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    return None

SLIDE_W_IN = 13.333
SLIDE_H_IN = 7.5

def px_to_in_x(px): return px * SLIDE_W_IN / 1280
def px_to_in_y(px): return px * SLIDE_H_IN / 720
def E(inch): return int(inch * 914400)
```

---

## 四、AI 生成 HTML Deck 的 Prompt 规范

GenSpark 的核心优势在于：**AI 直接生成 HTML**，不是生成 PPTX XML。HTML 对 AI 更自然，生成质量更高。

### Prompt 模板

```
你是一个专业的 PPT 设计师。请为以下主题生成一张 HTML 幻灯片（1280×720px）。

主题：{topic}
内容要点：
{bullet_points}

要求：
1. 使用以下 CDN（必须包含）：
   - Tailwind CSS: <script src="https://cdn.tailwindcss.com"></script>
   - FontAwesome 6: <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
   - Noto Sans SC: <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;700;900&display=swap" rel="stylesheet">

2. 画布规格：
   - 容器 id="slide-container"，width:1280px，height:720px
   - overflow:hidden，不能出现滚动条

3. 设计风格：
   - 背景色 #0f172a（深蓝）
   - 强调色 #38bdf8（天蓝）
   - 卡片背景 #1e293b，边框 #334155
   - 字体 'Noto Sans SC'

4. 必须包含：
   - 右上角和左下角各一个半透明装饰圆（position:absolute, opacity:0.05-0.1）
   - 标题左上角一条蓝色装饰横线（60px×4px）
   - FontAwesome 图标（每个功能点配一个 fa-solid 图标）
   - 内容卡片（圆角12px, 1px边框）

5. 输出：只输出完整 HTML 代码，不要解释。
```

---

## 五、完整 Pipeline：从 AI 生成到 PPTX 输出

```
┌─────────────────────────────────────────────────────────────────┐
│                   完整 Pipeline                                  │
│                                                                   │
│  内容输入 (Markdown/大纲/JSON)                                    │
│       ↓                                                          │
│  AI 生成 HTML deck                                               │
│  (含 Tailwind + FA + Noto Sans SC)                               │
│       ↓                                                          │
│  ┌────────────────────────────────────────────────────┐         │
│  │            转换方式选择                              │         │
│  │  A) 截图      B) 混合         C) 纯DOM解析          │         │
│  │  ↓            ↓               ↓                    │         │
│  │  PNG嵌入     截图+TextBox     Shape构建             │         │
│  │  快/完美     视觉好+可编辑    完全可编辑             │         │
│  └────────────────────────────────────────────────────┘         │
│       ↓                                                          │
│  PPTX 后处理：                                                    │
│  - 添加演讲者备注                                                 │
│  - 添加幻灯片编号                                                 │
│  - 设置切换动画                                                   │
│       ↓                                                          │
│  输出 .pptx 文件                                                  │
└─────────────────────────────────────────────────────────────────┘
```

### 实际调用示例

```python
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# 每个主题生成一张 HTML，然后转换
slides_html = []

for topic, bullets in content_plan:
    html = call_ai_generate_html(topic, bullets)  # 调用 Claude API
    slides_html.append(html)

# 混合模式转换
for html in slides_html:
    html_to_editable_pptx(html, prs)

prs.save('output.pptx')
```

---

## 六、各元素的 PPTX 最佳实现方案

| HTML 元素 | 推荐 PPTX 实现 | 备注 |
|-----------|---------------|------|
| `<i class="fa-solid fa-robot">` | 下载SVG→PNG→add_picture | 保留矢量质感 |
| `<canvas>` 折线图 | Playwright截图→add_picture | 或 python-pptx Chart |
| `<canvas>` 柱状图 | python-pptx 原生 BarChart | 完全可编辑 |
| CSS gradient 背景 | 纯色近似（PPTX不支持CSS渐变） | 取渐变中间色 |
| `filter: blur(120px)` 发光 | 低透明度大圆形（无blur） | 视觉效果近似 |
| Tailwind `bg-sky-400` | `RGBColor(0x38,0xBD,0xF8)` | 硬编码色值 |
| `border-radius: 12px` | `round_(shape, adj=8000)` | adj = 圆角/宽度*100000 |
| `text-shadow` | 不支持，省略 | PPTX无text-shadow |
| `<table>` | python-pptx Table 或矩形网格 | 原生支持 |
| `@keyframes` 动画 | PPTX 动画 Effect（可选） | 不影响静态演示 |
| SVG 图形 | add_picture（SVG作为图片） | 或解析SVG path |

---

## 七、关键优势总结：为什么 HTML → PPTX 优于直接生成 PPTX

| 维度 | HTML 先生成 | 直接生成 PPTX XML |
|------|------------|------------------|
| AI 理解难度 | ★ 简单（HTML是AI训练语料主体） | ★★★ 复杂（PPTX XML结构繁琐） |
| 视觉灵活性 | ★ 极高（CSS任意布局） | ★★ 受限（python-pptx API有限） |
| 图标支持 | ★ FA 2000+图标即插即用 | ★ 需要额外处理 |
| 数据可视化 | ★ Canvas/Chart.js 原生 | ★ python-pptx Chart有限 |
| 可编辑性 | ★★ 需转换（混合模式下OK） | ★ 完全可编辑 |
| 渲染一致性 | ★ 截图100%一致 | ★★ 跨Office版本有差异 |
| 生成速度 | ★★ 需要截图步骤（慢） | ★ 直接生成（快） |

**结论**：对于需要视觉丰富的技术演示，HTML先行 + 混合模式转换是最佳实践。
