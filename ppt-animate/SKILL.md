---
name: ppt-animate
description: |
  为 PowerPoint 演示文稿添加点击触发的转场和入场动画。支持 COM 自动化（需要安装 PowerPoint）和纯 XML 操作（无需 PowerPoint，跨平台）。包含 19 种转场效果和 22 种入场动画效果，提供多种风格模板。
  
  Add click-triggered transition and entrance animations to PowerPoint presentations. Supports both COM automation (requires PowerPoint installed) and pure XML manipulation (no PowerPoint needed, cross-platform). Includes 19 transition effects and 22 entrance animation effects with multiple style templates.
---

# PPT Animation Skill

## 功能特性 / Features

- **页内动画** / **In-slide Animations**: 每个内容形状添加 1 个 onClick 入场动画
  Add 1 onClick entrance animation to each content shape
- **幻灯片转场** / **Slide Transitions**: 所有幻灯片统一设置转场动画
  Apply uniform transition animations to all slides
- **双模式支持** / **Dual Mode Support**:
  - **COM 模式**: 需要安装 PowerPoint，通过 win32com 自动化操作
    COM mode: Requires PowerPoint installation, uses win32com automation
  - **XML 模式**: 纯 Python 操作 PPTX 内部 XML，无需 PowerPoint
    XML mode: Pure Python PPTX XML manipulation, no PowerPoint required
- **动画类型** / **Animation Types**: 19 种转场 + 22 种入场动画
  19 transitions + 22 entrance effects
- **风格模板** / **Style Templates**: 简洁商务、动态活力、优雅柔和、科技未来
  Business Simple, Dynamic Vitality, Elegant Soft, Tech Future

## 使用方法 / Usage

### 快速开始 / Quick Start

```python
from ppt_animate import animate_pptx

# 默认 COM 模式（需要 PowerPoint）/ Default COM mode (requires PowerPoint)
animate_pptx('input.pptx', 'output.pptx')

# XML 模式（无需 PowerPoint，跨平台）/ XML mode (no PowerPoint, cross-platform)
animate_pptx('input.pptx', 'output.pptx', method='xml')
```

### 自定义风格 / Custom Styles

```python
# 使用内置风格模板 / Use built-in style templates
animate_pptx('input.pptx', 'output.pptx', style='简洁商务')  # Business Simple
animate_pptx('input.pptx', 'output.pptx', style='动态活力')  # Dynamic Vitality
animate_pptx('input.pptx', 'output.pptx', style='优雅柔和')  # Elegant Soft
animate_pptx('input.pptx', 'output.pptx', style='科技未来')  # Tech Future
```

### 自定义转场 / Custom Transitions

```python
animate_pptx(
    'input.pptx',
    'output.pptx',
    transitions=['fade', 'zoom', 'fade'],
    method='xml'
)
```

## API 参考 / API Reference

### `animate_pptx(input_path, output_path=None, transitions=None, slide_groups=None, effects=None, method='com', style=None)`

为 PPT 文件添加动画。/ Add animations to PPT files.

**参数 / Parameters:**
- `input_path`: 输入 PPTX 文件路径 / Input PPTX file path
- `output_path`: 输出文件路径（默认：同目录 `_animated` 后缀）/ Output file path (default: `_animated` suffix in same directory)
- `transitions`: 转场效果列表（默认: `['fade'] * num_slides`）/ Transition effect list
- `slide_groups`: 形状分组列表（默认: 自动检测所有 TextBox 形状）/ Shape grouping list
- `effects`: 入场动画效果列表 / Entrance animation effect list
- `method`: `'com'`（需要 PowerPoint）或 `'xml'`（纯 Python）/ `'com'` (needs PowerPoint) or `'xml'` (pure Python)
- `style`: 风格模板，如 `'简洁商务'`、`'动态活力'` 等 / Style template

**返回 / Returns:** 输出文件路径 / Output file path

## 支持的动画效果 / Supported Effects

### 转场动画 (19 种) / Transitions (19)

| 名称 / Name | 说明 / Description | 名称 / Name | 说明 / Description |
|------------|-------------------|------------|-------------------|
| fade | 淡入淡出 / Fade in/out | wipe | 擦除 / Wipe |
| split | 分割 / Split | zoom | 缩放 / Zoom |
| push | 推出 / Push | cover | 覆盖 / Cover |
| dissolve | 溶解 / Dissolve | blind | 盲目 / Blind |
| checkerboard | 棋盘 / Checkerboard | coin | 硬币 / Coin |
| diamond | 菱形 / Diamond | random | 随机 / Random |
| ripple | 波纹 / Ripple | starburst | 星爆 / Starburst |
| sunbeam | 阳光 / Sunbeam | triangle | 三角形 / Triangle |
| wedge | 楔形 / Wedge | wheel | 轮子 / Wheel |
| zipper | 拉链 / Zipper | | |

### 入场动画 (22 种) / Entrance Effects (22)

| 名称 / Name | 说明 / Description | 名称 / Name | 说明 / Description |
|------------|-------------------|------------|-------------------|
| appear | 出现 / Appear | checkerboard | 棋盘 / Checkerboard |
| circle | 圆形展开 / Circle expand | box | 方框展开 / Box expand |
| spin | 旋转 / Spin | fly | 飞入 / Fly in |
| blend | 混合 / Blend | blur | 模糊 / Blur |
| compress | 压缩 / Compress | dissolve | 溶解 / Dissolve |
| explode | 爆炸 / Explode | fade | 淡入 / Fade |
| glow | 发光 / Glow | grow | 放大 / Grow |
| misty | 迷雾 / Misty | ripple | 波纹 / Ripple |
| reveal | 揭示 / Reveal | roll | 滚动 / Roll |
| shrink | 缩小 / Shrink | swizzle | 扭转 / Swizzle |
| teeter | 摇摆 / Teeter | typeWriter | 打字机 / Type writer |

## 注意事项 / Notes

### COM 模式 / COM Mode
- 需要安装 Microsoft PowerPoint / Requires Microsoft PowerPoint installation
- 代码已内置自动清理残留 PowerPoint 进程 / Auto cleanup for residual PowerPoint processes built-in
- 必须先处理 Slide 1，再处理 Slide 2-N（避免缓存问题）/ Process Slide 1 first, then Slide 2-N
- 每个形状只调用一次 AddEffect（避免重复动画）/ Call AddEffect only once per shape

### XML 模式 / XML Mode
- 无需 PowerPoint，纯 Python 实现 / No PowerPoint needed, pure Python implementation
- 适用于服务器环境或无 PowerPoint 场景 / Suitable for server environments or no-PowerPoint scenarios
- 转场和动画通过直接操作 PPTX 内部 XML 实现 / Transitions and animations via direct PPTX XML manipulation
- 已验证支持所有 19 种转场和 22 种入场效果 / Verified to support all 19 transitions and 22 entrance effects

## 依赖 / Dependencies

```bash
# 基本安装（仅 XML 模式）/ Basic install (XML mode only)
pip install ppt-animate

# 完整安装（包含 COM 模式）/ Full install (with COM mode)
pip install "ppt-animate[com]"

# 开发安装 / Development install
pip install "ppt-animate[dev]"
```

### 系统要求 / System Requirements

- Python 3.8+
- Windows: 需要安装 Microsoft PowerPoint（COM 模式）/ Requires Microsoft PowerPoint (COM mode)
- macOS/Linux: 仅支持 XML 模式 / XML mode only

## 项目结构 / Project Structure

```
ppt-animate/
├── scripts/
│   └── ppt_animate/
│       ├── __init__.py    # 包导出 / Package exports
│       ├── core.py        # 核心实现 / Core implementation
│       └── colors.py      # 颜色方案 / Color schemes
├── assets/                # 图标资源 / Icon assets
├── references/            # 技术文档 / Technical docs
├── test/                  # 测试代码 / Test code
├── SKILL.md               # Skill 定义 / Skill definition
├── README.md              # 项目说明 / Project readme
├── workflow.md            # 工作流规则 / Workflow rules
├── requirements.txt       # 依赖 / Dependencies
├── pyproject.toml         # 项目配置 / Project config
├── LICENSE                # MIT 许可证 / MIT License
└── .gitignore             # Git 忽略规则 / Git ignore rules
```

## License

[MIT](LICENSE) © ppt-animate contributors
