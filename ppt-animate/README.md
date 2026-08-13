# ppt-animate

<div align="center">

[![PyPI version](https://badge.fury.io/py/ppt-animate.svg)](https://badge.fury.io/py/ppt-animate)
[![Python Versions](https://img.shields.io/pypi/pyversions/ppt-animate.svg)](https://pypi.org/project/ppt-animate/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Tests](https://github.com/ppt-animate/ppt-animate/actions/workflows/test.yml/badge.svg)](https://github.com/ppt-animate/ppt-animate/actions)

**为 PowerPoint 演示文稿添加点击触发动画效果，支持 COM 自动化和纯 XML 操作两种模式。**

</div>

## 功能特性

- **页内动画**: 每个内容形状添加 1 个 onClick 入场动画（点击触发，不重复播放）
- **幻灯片转场**: 所有幻灯片统一设置转场动画
- **双模式支持**:
  - **COM 模式**: 需要安装 PowerPoint，通过 win32com 自动化操作
  - **XML 模式**: 纯 Python 操作 PPTX 内部 XML，无需 PowerPoint
- **动画类型**: 19 种转场 + 22 种入场动画
- **风格模板**: 简洁商务、动态活力、优雅柔和、科技未来

## 安装

```bash
# 基本安装（仅 XML 模式）
pip install ppt-animate

# 完整安装（包含 COM 模式）
pip install "ppt-animate[com]"

# 开发安装
pip install "ppt-animate[dev]"
```

### 系统要求

- Python 3.8+
- Windows: 需要安装 Microsoft PowerPoint（COM 模式）
- macOS/Linux: 仅支持 XML 模式

## 快速开始

### 为现有 PPT 添加动画

```python
from ppt_animate import animate_pptx

# 默认 COM 模式（需要 PowerPoint）
animate_pptx('input.pptx', 'output.pptx')

# XML 模式（无需 PowerPoint，跨平台）
animate_pptx('input.pptx', 'output.pptx', method='xml')
```

### 自定义动画风格

```python
# 使用内置风格模板
animate_pptx('input.pptx', 'output.pptx', style='简洁商务')
animate_pptx('input.pptx', 'output.pptx', style='动态活力')
animate_pptx('input.pptx', 'output.pptx', style='优雅柔和')
animate_pptx('input.pptx', 'output.pptx', style='科技未来')
```

### 自定义转场效果

```python
animate_pptx(
    'input.pptx',
    'output.pptx',
    transitions=['fade', 'zoom', 'fade'],
    method='xml'
)
```

## API 参考

### `animate_pptx(input_path, output_path=None, transitions=None, slide_groups=None, effects=None, method='com', style=None)`

为 PPT 文件添加动画。

**参数:**
- `input_path`: 输入 PPTX 文件路径
- `output_path`: 输出文件路径（默认：同目录 `_animated` 后缀）
- `transitions`: 转场效果列表（默认: `['fade'] * num_slides`）
- `slide_groups`: 形状分组列表（默认: 自动检测所有 TextBox 形状）
- `effects`: 入场动画效果列表
- `method`: `'com'`（需要 PowerPoint）或 `'xml'`（纯 Python）
- `style`: 风格模板，如 `'简洁商务'`、`'动态活力'` 等

**返回:** 输出文件路径

## 支持的动画效果

### 转场动画 (19 种)

| 名称 | 说明 | 名称 | 说明 |
|------|------|------|------|
| fade | 淡入淡出 | wipe | 擦除 |
| split | 分割 | zoom | 缩放 |
| push | 推出 | cover | 覆盖 |
| dissolve | 溶解 | blind | 盲目 |
| checkerboard | 棋盘 | coin | 硬币 |
| diamond | 菱形 | random | 随机 |
| ripple | 波纹 | starburst | 星爆 |
| sunbeam | 阳光 | triangle | 三角形 |
| wedge | 楔形 | wheel | 轮子 |
| zipper | 拉链 | | |

### 入场动画 (22 种)

| 名称 | 说明 | 名称 | 说明 |
|------|------|------|------|
| appear | 出现 | checkerboard | 棋盘 |
| circle | 圆形展开 | box | 方框展开 |
| spin | 旋转 | fly | 飞入 |
| blend | 混合 | blur | 模糊 |
| compress | 压缩 | dissolve | 溶解 |
| explode | 爆炸 | fade | 淡入 |
| glow | 发光 | grow | 放大 |
| misty | 迷雾 | ripple | 波纹 |
| reveal | 揭示 | roll | 滚动 |
| shrink | 缩小 | swizzle | 扭转 |
| teeter | 摇摆 | typeWriter | 打字机 |

## 项目结构

```
ppt-animate/
├── scripts/
│   └── ppt_animate/
│       ├── __init__.py    # 包导出
│       ├── core.py        # 核心实现
│       └── colors.py      # 颜色方案
├── assets/                # 图标资源
├── references/            # 技术文档
├── test/                  # 测试代码
├── SKILL.md               # Skill 定义
├── README.md              # 项目说明
├── workflow.md            # 工作流规则
├── requirements.txt       # 依赖
├── pyproject.toml         # 项目配置
├── LICENSE                # MIT 许可证
└── .gitignore             # Git 忽略规则
```

## 注意事项

### COM 模式
- 需要安装 Microsoft PowerPoint
- 代码已内置自动清理残留 PowerPoint 进程
- 必须先处理 Slide 1，再处理 Slide 2-N（避免缓存问题）
- 每个形状只调用一次 AddEffect（避免重复动画）

### XML 模式
- 无需 PowerPoint，纯 Python 实现
- 适用于服务器环境或无 PowerPoint 场景
- 转场和动画通过直接操作 PPTX 内部 XML 实现
- 已验证支持所有 19 种转场和 22 种入场效果

## 开发

```bash
# 克隆仓库
git clone https://github.com/ppt-animate/ppt-animate.git
cd ppt-animate

# 安装开发依赖
pip install -e ".[dev]"

# 运行测试
pytest test/
```

## 贡献

欢迎提交 Issue 和 Pull Request！

## License

[MIT](LICENSE) © ppt-animate contributors
