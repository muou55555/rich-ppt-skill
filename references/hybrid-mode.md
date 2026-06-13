# 混合模式：与 GordenPPTSkill 模板页合并

## 场景

希望用 GordenPPTSkill 的精美封面/结尾页，加上 Rich Mode 原创内容页。

## ⚠️ 已知问题与修复

**问题**：旧方式只复制 `p:spTree`（形状树），丢失了模板页的图片/背景 relationship，
导致含嵌入图片的模板幻灯片在 PowerPoint 中无法正常显示（空白或占位符断裂）。

**根本原因**：PPTX 中图片通过 `r:embed="rIdX"` 引用，而 relationship 定义在
`slide/slideN.xml.rels` 里。仅复制 XML 不复制 relationship 文件，图片引用悬空。

**修复**：使用 `copy_template_slide(prs, src_slide)` 代替手动复制 spTree，
该函数会：
1. 把 src_slide 所有非 layout relationship（图片、音频、视频等）注册到新幻灯片的 part
2. 记录 old rId → new rId 的映射，更新 XML 里的引用
3. 复制整个 `p:cSld`（含 `p:bg` 背景），而非仅 `p:spTree`

## 正确实现方式

```python
import copy
from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from scripts.components import new_presentation, save_pptx


def copy_template_slide(prs, src_slide):
    """
    将 src_slide（来自另一个 Presentation）完整复制为 prs 中的新幻灯片。
    正确传递所有图片/媒体 relationship，避免图像丢失或幻灯片无法正常显示。
    """
    blank = prs.slide_layouts[6]
    new_slide = prs.slides.add_slide(blank)
    new_part  = new_slide.part
    src_part  = src_slide.part

    # 1. 复制 relationship（跳过 slideLayout 自带的）
    rId_map = {}
    for rId, rel in list(src_part.rels.items()):
        if 'slideLayout' in rel.reltype:
            continue
        try:
            if rel.is_external:
                new_rId = new_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
            else:
                new_rId = new_part.relate_to(rel.target_part, rel.reltype)
            rId_map[rId] = new_rId
        except Exception as e:
            print(f"  [warn] skip rel {rId}: {e}")

    # 2. 深复制整个 cSld（含背景 p:bg 和形状 p:spTree）
    src_cSld = copy.deepcopy(src_slide._element.find(qn('p:cSld')))

    # 3. 更新 XML 里的 rId 引用
    if rId_map:
        xml_str = etree.tostring(src_cSld, encoding='unicode')
        for old_rId, new_rId in rId_map.items():
            if old_rId != new_rId:
                xml_str = xml_str.replace(f'"{old_rId}"', f'"{new_rId}"')
        src_cSld = etree.fromstring(xml_str)

    # 4. 替换 new_slide 的 cSld
    dst = new_slide._element
    old_cSld = dst.find(qn('p:cSld'))
    if old_cSld is not None:
        dst.remove(old_cSld)
    dst.append(src_cSld)

    return new_slide


# ── 完整混合模式示例 ─────────────────────────────────────────────────────────

# 1. 加载 GordenPPTSkill 模板
SRC = "/path/to/GordenPPTSkill/templates/YOUR_TEMPLATE/template.pptx"
src_prs = Presentation(SRC)
TMPL_IDXS = [0, 1, len(src_prs.slides) - 1]   # 封面、目录、结尾

# 2. 创建新演示文稿，生成 Rich Mode 内容页
prs = new_presentation()
# make_toc(prs, ...)
# make_cards(prs, ...)

# 3. 追加模板页（使用修复后的函数）
for idx in TMPL_IDXS:
    copy_template_slide(prs, src_prs.slides[idx])

# 4. 调整页面顺序（封面→目录→内容→结尾）
n_content = len(prs.slides) - len(TMPL_IDXS)
order = (
    [n_content, n_content + 1] +
    list(range(n_content)) +
    [n_content + 2]
)
sldIdLst = prs.slides._sldIdLst
slides_list = list(sldIdLst)
for el in slides_list: sldIdLst.remove(el)
for i in order: sldIdLst.append(slides_list[i])

# 5. 保存
save_pptx(prs, "hybrid_output.pptx")
```

## ❌ 旧方式（禁止使用）

```python
# 以下代码只复制了形状树，图片关系丢失 → 幻灯片无法正常显示
blank = prs.slide_layouts[6]
for xml_el in tmpl_xml:
    new_sl = prs.slides.add_slide(blank)
    dst = new_sl._element.find(qn('p:cSld')).find(qn('p:spTree'))
    src2 = xml_el.find(qn('p:cSld')).find(qn('p:spTree'))
    for ch in list(dst): dst.remove(ch)
    for ch in list(src2): dst.append(copy.deepcopy(ch))
# ↑ 这种方式已废弃，请使用 copy_template_slide()
```
