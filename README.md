# Gorden PPT Skill

> 史上最强原生PPT Skill，更适合中国宝宝。生成的效果不让你震惊，你来打我。
>
> ⚠️ **非商业使用**：本仓库及内置模板**仅供个人学习与研究**，禁止任何商业用途。
> 
> **支持定制私有化模板**：如果你想要Agent能按你公司的PPT模板来生成PPT，可以加我微信**duge360**定制。

## 交流群

扫码加入「PPT Skill 交流群」一起讨论 / 反馈问题 / 看新版本：

<p align="center">
  <img src="./assets/group-qr.jpg" alt="PPT Skill 交流群二维码" width="320" />
</p>

> ⚠️ 群二维码 7 天内有效；过期请来 [Issues](https://github.com/GordenSun/GordenPPTSkill/issues) 留言，我会贴新的。

## 几大特色
1、能生成信息密度高、排版复杂、看起来高大上的PPT，也支持生成简约、商务风格的PPT。适合国企、互联网大厂使用。

2、兼容所有模型，DeepSeek、小米Mimo、Claude、GPT均实测过，国产模型也能完成的非常好。

3、技能自动更新机制：如果我更新了可选用的PPT模板，使用技能时会自动更新技能。技能像软件一样可以更新

## 效果展示

以下是用本 Skill 生成的实际页面（信息密度高、排版复杂、商务质感）：

<p align="center">
  <img src="./assets/showcase-1.jpg" alt="效果展示 1 - 战略定位分析" width="720" />
</p>

<p align="center">
  <img src="./assets/showcase-2.jpg" alt="效果展示 2 - 价值冰山模型" width="720" />
</p>

<p align="center">
  <img src="./assets/showcase-3.jpg" alt="效果展示 3 - 产品化闭环" width="720" />
</p>

## 使用方法
极其简单，把这段提示词发给你的Agent即可：

安装这个Skill：https://github.com/GordenSun/GordenPPTSkill
然后使用这个Skill做一个复杂、豪华的PPT，读取本地XXX文件，来介绍XXX项目

## 谁要看这个

- 想给自己用的 AI 助手装一个"做 PPT"技能的人：**请读 [SKILL.md](./SKILL.md)**
- 想看本项目目录怎么组织：继续往下看本文件
- 想理解模板分类 / 推荐：**请读 [templates/INDEX.md](./templates/INDEX.md)**

## 快速开始（命令行）

```bash
# 1. 确认依赖
python3 -c "import pptx; print(pptx.__version__)"   # python-pptx 1.0+
soffice --version    # LibreOffice (仅渲染预览时需要)
which pdftoppm       # poppler   (仅渲染预览时需要)

# 2. 选定模板 + 写 edits.json，跑构建
python3 scripts/build_pptx.py \
    templates/minimal-business-summary/template.pptx \
    edits.json \
    out/final.pptx \
    --detail templates/minimal-business-summary/detail.json

# 3. (可选) 渲染最终预览图
python3 scripts/render_slides.py out/final.pptx out/preview --dpi 144
```

## 字体环境

模板大量使用 `微软雅黑`。如果你的机器没装它，配 `~/.config/fontconfig/fonts.conf` 加一条 alias：

```xml
<alias binding="strong">
  <family>微软雅黑</family>
  <accept>
    <family>WenQuanYi Micro Hei</family>
    <family>DengXian</family>
    <family>Noto Sans SC</family>
    <family>PingFang SC</family>
  </accept>
</alias>
```

(`brew install --cask font-noto-sans-sc`，或下载 WenQuanYi Micro Hei 放进 `~/Library/Fonts/` 并 `fc-cache -f`。)

## 目录速览

```
SKILL.md         # AI 入口文档
VERSION          # 1.0.0
CHANGELOG.md     # 人读变更
updates.json     # 机读变更
manifest.json    # 每文件版本 + sha256
scripts/         # 5 个面向使用者的脚本（build / render / update / manifest）
references/      # 编辑规则、Schema、工作流参考
templates/       # 17 个模板（每个 4 文件）
```

## 致谢与版权

- 本仓库没有PPT模板的版权
- **禁止任何二次分发 / 商业使用**
- 用到的开源工具：[LibreOffice](https://www.libreoffice.org/)、[python-pptx](https://python-pptx.readthedocs.io/)、[Poppler](https://poppler.freedesktop.org/)、[WenQuanYi Micro Hei](http://wenq.org/)
- 感谢 [LinuxDO](https://linux.do) 社区的支持

---

## Design Spec — GenSpark HTML Deck 设计规范（2024年分析）

本仓库 `design-spec/` 目录包含从 GenSpark AI PPT 生成产品中逆向分析总结的完整设计方法论。

### 文件说明

| 文件 | 内容 |
|------|------|
| [`design-spec/genspark-design-language.md`](design-spec/genspark-design-language.md) | 核心设计语言规范：配色、字体、布局、12个组件的 python-pptx 实现 |
| [`design-spec/slide-type-recipes.md`](design-spec/slide-type-recipes.md) | 9种页面类型配方：封面/目录/4卡片/6卡片/时间线/流程/统计/矩阵/章节页 |
| [`design-spec/pptx-component-library.py`](design-spec/pptx-component-library.py) | 可直接使用的 python-pptx 组件库（颜色常量、形状函数、布局助手） |
| [`design-spec/html-to-pptx-methodology.md`](design-spec/html-to-pptx-methodology.md) | HTML→PPTX 完整转换方法论（FontAwesome图标、3种转换模式、Playwright截图） |
| [`design-spec/dynamic-layout-generation.md`](design-spec/dynamic-layout-generation.md) | 动态布局生成：AI驱动、内容分类器、多布局Prompt模板、品质验证 |

### 核心洞察

**GenSpark 的技术路线**（HTML先行，非模板）：
1. AI 生成 HTML（含 Tailwind CSS + FontAwesome 6 + Noto Sans SC）
2. Playwright 截图 + DOM 解析文本坐标
3. 截图作背景 + 透明 TextBox 叠加 → 可编辑 PPTX

**核心 CDN 依赖**：
```html
<script src="https://cdn.tailwindcss.com"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;700;900&display=swap" rel="stylesheet">
```

**设计色板**（16 decks 统一使用）：
- 背景 `#0f172a` | 卡片 `#1e293b` | 边框 `#334155`
- 强调 `#38bdf8`（sky-400）| 主文字 `#f1f5f9` | 辅助 `#94a3b8`
