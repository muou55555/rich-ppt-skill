# PPT Builder Skill

> 让 AI 也能在不破坏排版的前提下，按用户需求挑选 / 编辑 / 输出一份 PPT。
>
> ⚠️ **非商业使用**：本仓库及内置模板**仅供个人学习与研究**，禁止任何商业用途。

## 这是什么

一套面向 AI 助手的"做 PPT"技能包。包含：

- **17 套精心整理的内置中文 PPT 模板**：每套都含原 .pptx、详细结构化说明（detail.json）、高度浓缩的风格简介（intro.md）、4 页 2×2 预览图（preview.png）
- **一套不破坏排版的编辑工具**：基于 python-pptx，按 JSON 指令选页 + 改文字，所有形状/字体/颜色保持不变
- **完整的工作流文档**：从"用户给你一段话" → "选模板" → "改文字" → "生成 PPT"全流程
- **版本管理 + 增量更新**：避免每次全量下载，知道哪些文件改了

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
