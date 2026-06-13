#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_pptx.py — Rich PPT 通用生成器
基于 components.py 提供常用幻灯片类型，可作为模板直接修改使用。

用法：
    python3 build_pptx.py --config config.json --output out.pptx
    python3 build_pptx.py --demo --output demo.pptx
"""
import argparse, json, os, sys
sys.path.insert(0, os.path.dirname(__file__))
from components import *

def _hex(c):
    return f"#{str(c)}"  # RGBColor -> "#RRGGBB"


# ── 预置幻灯片类型 ─────────────────────────────────────────────────────────────

def make_cover(prs, title: str, subtitle: str = '',
               meta: list = None, dark: bool = True):
    """
    封面页
    dark=True：深海蓝背景（正式技术汇报）
    dark=False：白底（商务/通用）
    meta: [(label, value), ...] 右侧信息卡片，最多3个
    """
    sl = blank_slide(prs)
    bg_color = NAVY if dark else WHITE
    fg_color = WHITE if dark else NAVY
    rect(sl, 0, 0, W, int(H * 0.62), bg_color)
    rect(sl, 0, int(H * 0.62), W, int(H * 0.38), WHITE)
    rect(sl, int(W * 0.024), int(H * 0.10), int(W * 0.006), int(H * 0.50), GOLD)
    box(sl, int(W * 0.05), int(H * 0.12), int(W * 0.78), int(H * 0.18),
        title, size=32, bold=False, color=fg_color)
    if subtitle:
        box(sl, int(W * 0.05), int(H * 0.31), int(W * 0.72), int(H * 0.08),
            subtitle, size=14,
            color=RGBColor(0xBB, 0xCC, 0xFF) if dark else GRAY)
    line_h(sl, int(W * 0.05), int(H * 0.43), int(W * 0.50), GOLD)
    if meta:
        for i, (label, val) in enumerate(meta[:3]):
            bx = int(W * 0.72) + i * int(W * 0.092)
            rect(sl, bx, int(H * 0.24), int(W * 0.085), int(H * 0.16), MED, r=10000)
            box(sl, bx, int(H * 0.25), int(W * 0.085), int(H * 0.06),
                label, size=9, color=RGBColor(0xAA, 0xBB, 0xFF) if dark else GRAY,
                align=PP_ALIGN.CENTER)
            box(sl, bx, int(H * 0.30), int(W * 0.085), int(H * 0.07),
                val, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    return sl


def make_toc(prs, chapters: list):
    """
    目录页
    chapters: [{'num':'01','title':'...','desc':'...','icon':'lightbulb','color':NAVY}, ...]
    """
    sl = blank_slide(prs)
    slide_bg(sl)
    title_block(sl, '目录', f'共 {len(chapters)} 章节')
    icon_sz = int(H * 0.058)
    for i, ch in enumerate(chapters):
        col = i % 2
        row = i // 2
        bx = int(W * 0.03) + col * int(W * 0.50)
        by = int(H * 0.215) + row * int(H * 0.155)
        bw = int(W * 0.47)
        bh = int(H * 0.14)
        clr = ch.get('color', ACCENTS[i % len(ACCENTS)])
        ico = ch.get('icon', 'settings')
        rect(sl, bx, by, bw, bh, WHITE, r=12000)
        rect(sl, bx, by, int(W * 0.006), bh, clr)
        bdr = sl.shapes.add_shape(1, Emu(bx), Emu(by), Emu(bw), Emu(bh))
        bdr.fill.background()
        bdr.line.color.rgb = LINE
        bdr.line.width = Pt(0.5)
        ix = bx + int(W * 0.015)
        iy = by + (bh - icon_sz) // 2
        add_icon(sl, ico, ix, iy, icon_sz, _hex(clr))
        box(sl, ix + icon_sz + int(W * 0.010), by + int(H * 0.016),
            int(W * 0.03), int(H * 0.055),
            ch.get('num', str(i+1)), size=13, bold=True, color=clr)
        box(sl, ix + icon_sz + int(W * 0.042), by + int(H * 0.016),
            bw - int(W * 0.16), int(H * 0.055),
            ch['title'], size=12, bold=True, color=NAVY)
        box(sl, ix + icon_sz + int(W * 0.010), by + int(H * 0.078),
            bw - int(W * 0.15), int(H * 0.055),
            ch.get('desc', ''), size=10, color=GRAY)
    return sl


def make_cards(prs, title: str, subtitle: str, items: list, cols: int = 2):
    """
    圆角卡片布局页
    items: [{'icon':'...','color':NAVY,'title':'...','body':'...','num':1}, ...]
    cols: 2 或 3
    """
    sl = blank_slide(prs)
    slide_bg(sl)
    title_block(sl, title, subtitle)
    n = len(items)
    rows = (n + cols - 1) // cols
    bw = int(W * (0.94 / cols)) - int(W * 0.01)
    bh = int(H * (0.76 / rows)) - int(H * 0.01)
    for i, item in enumerate(items):
        col = i % cols
        row = i // cols
        bx = int(W * 0.03) + col * (bw + int(W * 0.01))
        by = int(H * 0.215) + row * (bh + int(H * 0.01))
        accent = item.get('color', ACCENTS[i % len(ACCENTS)])
        rcard(sl, bx, by, bw, bh,
              item['title'],
              item.get('body', ''),
              accent=accent,
              num=item.get('num'),
              icon=item.get('icon'))
    return sl


def make_flow(prs, title: str, subtitle: str, steps: list):
    """
    箭头流程页
    steps: [{'label':'步骤名','sub':'说明'}, ...]
    """
    sl = blank_slide(prs)
    slide_bg(sl)
    title_block(sl, title, subtitle)
    chevron_flow(sl,
                 [(s['label'], s.get('sub', '')) for s in steps],
                 int(H * 0.32), int(H * 0.18))
    return sl


def make_table(prs, title: str, subtitle: str,
               headers: list, rows: list):
    """表格布局页"""
    sl = blank_slide(prs)
    slide_bg(sl)
    title_block(sl, title, subtitle)
    htable(sl, headers, rows,
           int(W * 0.03), int(H * 0.215), int(W * 0.94), int(H * 0.72))
    return sl


def make_timeline(prs, title: str, subtitle: str, events: list):
    """
    时序/事件列表页
    events: [{'label':'事件名','icon':'...','color':NAVY,'desc':'说明'}, ...]
    """
    sl = blank_slide(prs)
    slide_bg(sl)
    title_block(sl, title, subtitle)
    sh = int(H * 0.109)
    icon_sz = int(H * 0.060)
    lx = int(W * 0.03)
    for i, ev in enumerate(events):
        clr = ev.get('color', ACCENTS[i % len(ACCENTS)])
        ico = ev.get('icon', 'arrow-right')
        sy = int(H * 0.215) + i * (sh + int(H * 0.006))
        rect(sl, lx, sy, int(W * 0.94), sh, WHITE, r=8000)
        rect(sl, lx, sy, int(W * 0.006), sh, clr)
        bdr = sl.shapes.add_shape(1, Emu(lx), Emu(sy), Emu(int(W * 0.94)), Emu(sh))
        bdr.fill.background()
        bdr.line.color.rgb = LINE
        bdr.line.width = Pt(0.5)
        add_icon(sl, ico, lx + int(W * 0.012), sy + (sh - icon_sz) // 2, icon_sz, _hex(clr))
        tx = lx + int(W * 0.012) + icon_sz + int(W * 0.008)
        box(sl, tx, sy + int(H * 0.010), int(W * 0.15), int(H * 0.048),
            ev['label'], size=11, bold=True, color=clr)
        box(sl, tx + int(W * 0.155), sy + int(H * 0.012), int(W * 0.72), int(H * 0.082),
            ev.get('desc', ''), size=10, color=DARK)
    return sl


# ── CLI ───────────────────────────────────────────────────────────────────────

def _demo(output_path: str):
    """生成演示 PPT（展示所有组件）"""
    prs = new_presentation()
    make_cover(prs, 'Rich PPT Skill 演示', '组件库 · 图标系统 · 样式规范',
               meta=[('版本', '1.0'), ('组件', '8种'), ('图标', '50+')])
    make_toc(prs, [
        {'num':'01','title':'封面与目录','desc':'cover + toc 组件',
         'icon':'lightbulb','color':NAVY},
        {'num':'02','title':'圆角卡片','desc':'rcard 2×2/3×2 布局',
         'icon':'layers','color':SKY},
        {'num':'03','title':'箭头流程','desc':'chevron_flow 横排',
         'icon':'refresh','color':GREEN},
        {'num':'04','title':'高亮表格','desc':'htable 交替行色',
         'icon':'database','color':GOLD},
    ])
    make_cards(prs, '核心设计原则', '单循环 · 模型驱动 · Unix 哲学', [
        {'icon':'refresh', 'color':NAVY, 'title':'单一主循环',
         'body':'while True 扁平循环，让模型自己决策下一步，无复杂状态机。'},
        {'icon':'agent',   'color':SKY,  'title':'模型即调度器',
         'body':'工具顺序、执行时机、停止条件全部由模型决策，工程侧只做权限与恢复。'},
        {'icon':'terminal','color':GREEN,'title':'Unix 哲学',
         'body':'可管道化、可组合进 CI/CD，单一职责，运行在终端。'},
        {'icon':'settings','color':GOLD, 'title':'复杂度在边界',
         'body':'30行能写基础循环，生产级 1800+ 行——复杂性在"循环失败时如何优雅恢复"。'},
    ])
    make_flow(prs, 'Agent 主循环六步', '异步生成器驱动的扁平循环', [
        {'label':'① 组装请求', 'sub':'system + tools + history'},
        {'label':'② API 调用',  'sub':'SSE 流式'},
        {'label':'③ 解析响应',  'sub':'text or tool_use'},
        {'label':'④ 执行工具',  'sub':'权限检查 + Hook'},
        {'label':'⑤ 回传结果',  'sub':'tool_result'},
        {'label':'⑥ 继续循环',  'sub':'continue'},
    ])
    make_table(prs, '工具系统分类', '五大类 ~20 个内置工具', 
               ['类别', '代表工具', '要点'],
               [['文件操作', 'Read/Write/Edit', '强制先 Read 后 Edit'],
                ['搜索', 'Glob/Grep', 'ripgrep，返回结构化结果'],
                ['执行', 'Bash', '持久 Shell，输出超 30K 截断'],
                ['网络', 'WebSearch/WebFetch', '内容先摘要再填入上下文'],
                ['编排', 'Task/TodoWrite', '子代理派生，任务清单']])
    save_pptx(prs, output_path)


def main():
    parser = argparse.ArgumentParser(description='Rich PPT Builder')
    parser.add_argument('--config',  help='JSON 配置文件路径')
    parser.add_argument('--output',  default='/tmp/rich_ppt_output.pptx', help='输出路径')
    parser.add_argument('--demo',    action='store_true', help='生成演示 PPT')
    args = parser.parse_args()

    if args.demo:
        _demo(args.output)
        return

    if args.config:
        with open(args.config) as f:
            cfg = json.load(f)
        prs = new_presentation()
        for slide_cfg in cfg.get('slides', []):
            st = slide_cfg.get('type', 'cards')
            if st == 'cover':
                make_cover(prs, **{k: v for k, v in slide_cfg.items() if k != 'type'})
            elif st == 'toc':
                make_toc(prs, **{k: v for k, v in slide_cfg.items() if k != 'type'})
            elif st == 'cards':
                make_cards(prs, **{k: v for k, v in slide_cfg.items() if k != 'type'})
            elif st == 'flow':
                make_flow(prs, **{k: v for k, v in slide_cfg.items() if k != 'type'})
            elif st == 'table':
                make_table(prs, **{k: v for k, v in slide_cfg.items() if k != 'type'})
            elif st == 'timeline':
                make_timeline(prs, **{k: v for k, v in slide_cfg.items() if k != 'type'})
        save_pptx(prs, args.output)
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
