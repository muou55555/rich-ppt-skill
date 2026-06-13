# Rich PPT 生成 Pipeline 说明

## 整体架构

```
用户输入（文档/大纲）
    ↓
Claude 解析内容结构
    ↓
规划幻灯片类型和数量
    ↓
生成 Python 脚本（调用 components.py）
    ↓
运行脚本 → .pptx
    ↓
present_files 交付
```

## 依赖安装

```bash
# Python 依赖
pip install python-pptx cairosvg --break-system-packages -q

# 图标库（MIT 开源，扁平填充风格）
npm install --prefix /tmp/phosphor @phosphor-icons/core

# GordenPPTSkill 模板（混合模式用）
git clone --depth 1 https://github.com/GordenSun/GordenPPTSkill.git /tmp/GordenPPTSkill
```

## 图标获取方案对比

| 方案 | 优点 | 缺点 | 推荐度 |
|---|---|---|---|
| Phosphor Icons (npm) | 高质量，MIT，扁平填充，本地无网络 | 需 npm 安装 | ⭐⭐⭐⭐⭐ |
| Heroicons (内嵌 SVG) | 无依赖 | Stroke 风格，非扁平 | ⭐⭐⭐⭐ |
| iconfont.cn | 图标量大，扁平风格 | 需登录下载，无法自动化 | ⭐⭐（需手动） |
| 手写 flat SVG | 无依赖 | 路径不精确，易变形 | ⭐⭐ |

**结论**：优先 Phosphor Icons（npm 安装）；若 npm 不可用则回退 Heroicons 内嵌 SVG。

## 关键技术点

### 1. EMU 坐标系
python-pptx 使用 EMU（English Metric Units），1 英寸 = 914400 EMU。
幻灯片：W=12192000, H=6858000（宽屏 16:9）。

### 2. 圆角矩形
通过修改 XML 中的 `a:avLst/a:gd` 元素实现：
```python
gd.set('fmla', f'val {radius}')  # radius 单位：1/100 度
```

### 3. Phosphor SVG 着色
Phosphor SVG 使用 `currentColor`，直接替换为目标颜色：
```python
svg = svg.replace("currentColor", "#00329D")
```

### 4. 孤儿关系清理
操作 `_sldIdLst` 重排幻灯片后必须调用：
```python
purge_orphans(prs)
```

### 5. 模板混合
从 GordenPPTSkill 提取模板页 XML，复制到新演示文稿。
需要 `copy.deepcopy()` 以避免共享 XML 元素。
