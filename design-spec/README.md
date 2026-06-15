# Rich PPT Skill — 专业技术PPT生成规范项目

基于 GenSpark HTML Deck 设计语言深度分析，总结了生成高质量技术演示文稿（PPTX）的完整方法论。

## 目录结构

```
rich-ppt-skill/
├── README.md                          ← 本文件
├── design-spec/
│   ├── genspark-design-language.md   ← 完整设计语言规范（核心文档）
│   ├── slide-type-recipes.md         ← 各类幻灯片的实现配方手册
│   └── pptx-component-library.py    ← python-pptx 组件库（可直接使用）
└── examples/
    └── (生成的示例PPT文件)
```

## 核心设计语言

### 配色（暗色主题）
- 主背景 `#0f172a` | 卡片 `#1e293b` | 边框 `#334155`
- 强调色 `#38bdf8`（sky-400）
- 主文字 `#f1f5f9` | 辅助文字 `#94a3b8`

### 画布规格
- 尺寸：1280×720px = 13.333" × 7.5"（python-pptx 标准 16:9）
- 字体：Noto Sans SC

### 核心组件
1. **装饰线（Deco Line）**：标题上方60px蓝色横线
2. **发光圆（BG Glow）**：右上+左下，模糊，5%透明度
3. **功能卡片（Feature Card）**：圆角12px + 图标盒 + 标题 + 描述
4. **胶囊标签（Badge）**：圆角50000，用于分类标注
5. **装饰圆（Decor Circle）**：角落装饰，增加层次感

## 快速使用

```python
from design-spec.pptx-component-library import *

prs = Presentation()
prs.slide_width = E(13.333)
prs.slide_height = E(7.5)

# 创建一张带Header的内容页
slide = make_content_slide(prs, '核心特性', '架构深度解析', badge='技术架构')

# 添加4个功能卡片
cards = [
    ('🤖', C['accent'], RGBColor(0x0A,0x1F,0x3A), '多智能体', '并行subagent协同处理'),
    ('🔀', C['accent'], RGBColor(0x0A,0x1F,0x3A), '工具集成', '30+内置工具无缝调用'),
    ('🛡️', C['accent'], RGBColor(0x0A,0x1F,0x3A), '安全模型', '最小权限原则'),
    ('⚡', C['accent'], RGBColor(0x0A,0x1F,0x3A), '流式响应', '低延迟实时输出'),
]
layout_card_grid(slide, cards, cols=2)

prs.save('output.pptx')
```

## 关键设计原则

1. **暗色优先**：所有页面使用深色背景，与技术/科技感一致
2. **单一强调色**：只用 `#38bdf8`（sky blue）作强调，避免彩虹色
3. **层次清晰**：标题27pt → 卡片标题11pt → 正文9pt
4. **留白充足**：每边留白 ≥ 0.4"，卡片间距 ≥ 0.17"
5. **圆角一致**：卡片 adj=8000，Badge adj=50000，图标盒 adj=8000
6. **内容密度**：每张幻灯片 4-6 个信息单元，不超过7个

## 设计来源

- **GenSpark AI**（genspark.ai）：分析其HTML Deck生成产品的16个实际案例
- 技术栈：Tailwind CSS + FontAwesome 6 + Noto Sans SC + Custom CSS
- 画布渲染：1280×720px HTML→PNG→PPTX 转换流程
