# 图标参考手册

图标来源：**Phosphor Icons fill**（MIT 开源，扁平实色，视觉等同 iconfont 面性图标）
npm: `@phosphor-icons/core`

## 使用方法

```python
from scripts.icons import get_icon
png_path = get_icon("lightbulb", "#00329D", 64)
slide.shapes.add_picture(png_path, Emu(x), Emu(y), Emu(size_emu), Emu(size_emu))
```

## 完整图标列表

| 语义名 | Phosphor 文件 | 适用场景 |
|---|---|---|
| lightbulb | lightbulb-filament-fill | 设计哲学、创意、洞察 |
| terminal | terminal-fill | 命令行、代码执行 |
| refresh | arrow-clockwise-fill | 循环、刷新、更新 |
| tool | wrench-fill | 工具、配置、维护 |
| layers | stack-fill | 图层、堆叠、上下文 |
| puzzle | puzzle-piece-fill | 扩展、插件、集成 |
| shield | shield-check-fill | 安全、权限、防护 |
| cpu | cpu-fill | 模型、芯片、计算 |
| copy | copy-fill | 复制、复刻、克隆 |
| book | book-bookmark-fill | 文档、参考、学习 |
| file-code | file-code-fill | 代码文件、脚本 |
| search | magnifying-glass-fill | 搜索、探索、查找 |
| command | terminal-window-fill | 命令、执行、终端窗口 |
| globe | globe-fill | 网络、全球、Web |
| git-branch | git-branch-fill | 版本控制、分支 |
| lock | lock-fill | 加密、权限、安全锁 |
| database | database-fill | 数据库、持久化、存储 |
| dollar | currency-dollar-fill | 成本、定价、费用 |
| check-circle | check-circle-fill | 完成、验证、成功 |
| hook | webhooks-logo-fill | Hooks、钩子、事件 |
| memory | brain-fill | 内存、记忆、知识 |
| arrow-right | arrow-right-fill | 箭头、方向、导航 |
| server | desktop-fill | 服务器、桌面、主机 |
| agent | robot-fill | 智能体、AI、自动化 |
| settings | gear-six-fill | 设置、配置、齿轮 |
| network | network-fill | 网络、连接、拓扑 |
| chart | chart-bar-fill | 图表、统计、数据 |
| user | user-fill | 用户、个人 |
| team | users-fill | 团队、组织 |
| cloud | cloud-fill | 云服务、部署 |
| code | code-fill | 代码、编程 |
| warning | warning-fill | 警告、注意 |
| rocket | rocket-fill | 发布、上线、加速 |
| lightning | lightning-fill | 快速、性能、闪电 |

## 颜色建议

| 语境 | 推荐颜色 |
|---|---|
| 主要功能 | `#00329D`（NAVY） |
| 辅助功能 | `#4472C4`（SKY） |
| 成功/完成 | `#1E8C55`（GREEN） |
| 警告/注意 | `#EDAD1A`（GOLD） |
| 危险/错误 | `#C0392B`（RED2） |
| 白色背景上 | `#00329D` 或各 accent 色 |
| 深色背景上 | `#FFFFFF` |
