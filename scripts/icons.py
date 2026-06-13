#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
icons.py — Phosphor Icons fill 风格（扁平实色，MIT 开源）
视觉等同 iconfont 面性图标。
依赖：cairosvg, @phosphor-icons/core（npm）

用法：
    from scripts.icons import get_icon
    png_path = get_icon("lightbulb", "#00329D", 64)
    slide.shapes.add_picture(png_path, Emu(x), Emu(y), Emu(w), Emu(h))
"""
import os, sys

# cairosvg 可能安装在非标准路径
for _p in [
    '/sessions/youthful-sharp-gauss/.local/lib/python3.10/site-packages',
    os.path.expanduser('~/.local/lib/python3.10/site-packages'),
    os.path.expanduser('~/.local/lib/python3.11/site-packages'),
]:
    if os.path.isdir(_p) and _p not in sys.path:
        sys.path.insert(0, _p)

try:
    import cairosvg
    _CAIRO_OK = True
except ImportError:
    _CAIRO_OK = False

# Phosphor Icons fill SVG 目录（npm install @phosphor-icons/core）
_PHOSPHOR_CANDIDATES = [
    "/tmp/phosphor/node_modules/@phosphor-icons/core/assets/fill",
    os.path.expanduser("~/node_modules/@phosphor-icons/core/assets/fill"),
    "/usr/local/lib/node_modules/@phosphor-icons/core/assets/fill",
]

def _find_phosphor():
    for p in _PHOSPHOR_CANDIDATES:
        if os.path.isdir(p):
            return p
    return None

PHOSPHOR_FILL_DIR = _find_phosphor()

ICON_CACHE_DIR = "/tmp/rich_ppt_icons"
os.makedirs(ICON_CACHE_DIR, exist_ok=True)

# 名称映射：语义名 → Phosphor 文件名（不含 -fill.svg）
MAP = {
    "lightbulb":     "lightbulb-filament",
    "terminal":      "terminal",
    "refresh":       "arrow-clockwise",
    "tool":          "wrench",
    "layers":        "stack",
    "puzzle":        "puzzle-piece",
    "shield":        "shield-check",
    "cpu":           "cpu",
    "copy":          "copy",
    "book":          "book-bookmark",
    "file-code":     "file-code",
    "search":        "magnifying-glass",
    "command":       "terminal-window",
    "globe":         "globe",
    "git-branch":    "git-branch",
    "lock":          "lock",
    "database":      "database",
    "dollar":        "currency-dollar",
    "check-circle":  "check-circle",
    "hook":          "webhooks-logo",
    "memory":        "brain",
    "arrow-right":   "arrow-right",
    "server":        "desktop",
    "agent":         "robot",
    "settings":      "gear-six",
    "network":       "network",
    # 扩展图标
    "chart":         "chart-bar",
    "user":          "user",
    "team":          "users",
    "calendar":      "calendar",
    "clock":         "clock",
    "star":          "star",
    "warning":       "warning",
    "info":          "info",
    "cloud":         "cloud",
    "code":          "code",
    "home":          "house",
    "folder":        "folder",
    "mail":          "envelope",
    "phone":         "phone",
    "link":          "link",
    "play":          "play-circle",
    "pause":         "pause-circle",
    "download":      "download-simple",
    "upload":        "upload-simple",
    "edit":          "pencil-simple",
    "delete":        "trash",
    "add":           "plus-circle",
    "close":         "x-circle",
    "flag":          "flag",
    "tag":           "tag",
    "key":           "key",
    "graph":         "graph",
    "lightning":     "lightning",
    "rocket":        "rocket",
}

# 备用内嵌 SVG（Heroicons stroke 风格，当 Phosphor 不可用时）
_FALLBACK = {
    "lightbulb": '<svg viewBox="0 0 24 24" fill="none" stroke="{c}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18h6M10 22h4M12 2a7 7 0 0 1 7 7c0 2.7-1.5 5-3.7 6.3V17H8.7V15.3C6.5 14 5 11.7 5 9a7 7 0 0 1 7-7z"/></svg>',
    "terminal":  '<svg viewBox="0 0 24 24" fill="none" stroke="{c}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="4 17 10 11 4 5"/><line x1="12" y1="19" x2="20" y2="19"/></svg>',
    "settings":  '<svg viewBox="0 0 24 24" fill="none" stroke="{c}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>',
}

def get_icon(name: str, color: str = "#00329D", size: int = 64) -> str:
    """
    返回图标 PNG 文件路径（带缓存）。

    Args:
        name:  图标语义名，见 MAP 字典
        color: 十六进制颜色，如 "#00329D"
        size:  输出像素尺寸（正方形）

    Returns:
        PNG 文件绝对路径
    """
    safe_color = color.replace("#", "").upper()
    cache_path = os.path.join(ICON_CACHE_DIR, f"{name}_{safe_color}_{size}.png")
    if os.path.exists(cache_path):
        return cache_path

    svg_text = None

    # 优先：Phosphor fill SVG
    if PHOSPHOR_FILL_DIR:
        phosphor_name = MAP.get(name, "gear-six")
        svg_file = os.path.join(PHOSPHOR_FILL_DIR, f"{phosphor_name}-fill.svg")
        if os.path.exists(svg_file):
            with open(svg_file) as f:
                svg_text = f.read()
            svg_text = svg_text.replace("currentColor", color)
            if "fill=" not in svg_text:
                svg_text = svg_text.replace("<svg ", f'<svg fill="{color}" ', 1)

    # 备用：内嵌 stroke SVG
    if svg_text is None:
        fallback_key = name if name in _FALLBACK else "settings"
        svg_text = _FALLBACK[fallback_key].replace("{c}", color)

    if not _CAIRO_OK:
        raise RuntimeError(
            "cairosvg 未安装。请运行: pip install cairosvg --break-system-packages"
        )

    cairosvg.svg2png(
        bytestring=svg_text.encode("utf-8"),
        write_to=cache_path,
        output_width=size,
        output_height=size,
    )
    return cache_path


def list_icons() -> list:
    """返回所有可用图标名称列表"""
    return list(MAP.keys())


if __name__ == "__main__":
    print("Testing icon generation...")
    if not PHOSPHOR_FILL_DIR:
        print("⚠ Phosphor not found. Run: npm install --prefix /tmp/phosphor @phosphor-icons/core")
    for name in list(MAP.keys())[:5]:
        try:
            p = get_icon(name, "#00329D", 64)
            print(f"  ✓ {name:20s} → {p}")
        except Exception as e:
            print(f"  ✗ {name}: {e}")
    print("Done.")
