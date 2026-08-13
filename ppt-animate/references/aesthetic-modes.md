# 动画美学模式规范

> 基于 huashu-design 的 Anthropic 级动画设计哲学，为 PPT 动画添加审美层次。
> **核心信念：动画是物理学，不是 CSS transition。**

---

## 一、动画三层架构

### 层 1：叙事节奏层（必须遵循）
`
Slow → Fast → Boom → Stop
`

| 阶段 | 时长占比 | 作用 | 实现方式 |
|------|----------|------|----------|
| **Slow** | 30% | 建立预期 | 元素从远处缓慢靠近（duration 0.6-0.8s） |
| **Fast** | 40% | 建立动量 | 加速进入位置（easing: expoOut） |
| **Boom** | 20% | 情感峰值 | 关键结果前的停顿（delay 0.3-0.5s） |
| **Stop** | 10% | 决定感 | 戛然而止，hold 住最后一帧 |

### 层 2：运动曲线层（优先级）
`
expoOut > overshoot > spring > easeOutCubic > linear
`

| 曲线名 | 对应 XML | 适用场景 |
|--------|----------|----------|
| xpoOut | calcMode="spline" keySplines="0.19 1 0.22 1" | 默认入场（物体落地感） |
| overshoot | calcMode="spline" keySplines="0.34 1.56 0.64 1" | 按钮/卡片弹出 |
| spring | calcMode="spline" keySplines="0.34 1.56 0.64 1" | 弹性效果 |
| aseOutCubic | calcMode="spline" keySplines="0.33 1 0.68 1" | 平滑过渡 |

### 层 3：时间分配层
`
总时长 = 单元素 × N + Stagger × (N-1)
`

| 元素数 | 建议单时长 | Stagger | 总时长 |
|--------|------------|---------|--------|
| 1-3 | 0.4-0.6s | 0.05s | 0.5-1.0s |
| 4-6 | 0.3-0.5s | 0.08s | 0.8-1.5s |
| 7-10 | 0.2-0.4s | 0.10s | 1.0-2.0s |
| >10 | 0.15-0.3s | 0.12s | 1.5-3.0s |

---

## 二、动画类型选择指南

### 2.1 入场动画适配表

| 内容类型 | 推荐动画 | 时长 | Easing | 说明 |
|----------|----------|------|--------|------|
| **封面标题** | zoom (41) + fade (5) | 0.8s | overshoot | 形变入场有冲击力 |
| **要点列表** | fly (4) + fade (5) | 0.4s | expoOut | 错开 stagger 0.08s |
| **图表数据** | blur (28) + grow (53) | 0.6s | spring(0.6) | 模糊→清晰有科技感 |
| **金句断言** | fade (5) + grow (53) | 0.5s | easeOutCubic | 居中浮现庄重感 |
| **流程图** | roll (66) + wipe (7) | 0.5s | expoOut | 滚动展开叙事感 |
| **团队介绍** | converge (无) → split (13) | 0.7s | overshoot | 从四周汇聚 |

### 2.2 转场动画情感匹配

| 转场类型 | 情感 | 适用场景 |
|----------|------|----------|
| **fade** | 从容、优雅 | 章节过渡、平稳叙事 |
| **push** | 向前、果断 | 流程推进、步骤递进 |
| **split** | 冲突、张力 | 对比展示、问题/方案 |
| **zoom** | 冲击、聚焦 | 重点强调、视角切换 |
| **wipe** | 揭示、透明 | 数据展示、揭秘环节 |
| **cover** | 覆盖、重置 | 场景切换、新篇章 |

---

## 三、三种产品性格配方

### 配方 A：轻快型（Fast-Punchy）
`yaml
适用场景: 初创公司、产品 launch、科技媒体
节奏特征: Fast-Punchy
入场动画: 
  - duration: 0.3-0.4s
  - stagger: 0.06s
  - easing: expoOut
转场动画:
  - 推荐: push, zoom
  - 避免: fade（太慢）
视觉感受: 节奏明快，信息密集，有冲击力
`

### 配方 B：平衡型（Medium-Smooth）⭐ 默认推荐
`yaml
适用场景: 企业汇报、商业路演、咨询报告
节奏特征: Medium-Smooth
入场动画:
  - duration: 0.4-0.6s
  - stagger: 0.08s
  - easing: spring(0.6)
转场动画:
  - 推荐: fade, wipe, split
  - 混合使用增加层次
视觉感受: 从容不迫，专业稳重，有呼吸感
`

### 配方 C：慢重叙事型（Slow-Epic）
`yaml
适用场景: 品牌宣传片、发布会 keynote、重大发布
节奏特征: Slow-Epic
入场动画:
  - duration: 0.6-0.8s
  - stagger: 0.12s
  - easing: overshoot
转场动画:
  - 推荐: zoom, cover, split
  - 关键节点用慢速强调
视觉感受: 庄重宏大，每个元素都有分量感
`

---

## 四、动画避坑清单（16 条铁律）

### ❌ 绝对禁止

| 问题 | 错误表现 | 正确做法 |
|------|----------|----------|
| **全 fade** | 所有入场都用 opacity | 混合 fly + zoom + blur |
| **匀速运动** | 	ransition: all 0.3s ease | 用 xpoOut 或 overshoot |
| **同时淡入** | 所有元素同时出现 | 加 stagger 0.08-0.12s |
| **Logo 淡入** | 无叙事淡入 | Morph 形变或 converge |
| **直线鼠标** | 直线移动轨迹 | 贝塞尔弧线 + Perlin Noise |
| **打字单字** | setInterval 单字蹦出 | Chunk Reveal（词组为单位） |
| **无悬停** | 关键结果立即切换 | 结果前 hold 0.5s |
| **焦点切换** | 只改 opacity | 同时加 blur 景深效果 |
| **纯黑/纯白底** | #000000 / #FFFFFF | 带色温的中性色（L 0.92-0.98） |
| **均匀节奏** | 所有动画同样快 | Slow-Fast-Boom-Stop 呼吸感 |
| **Fade out 收尾** | 渐弱结束 | 戛然而止，hold 最后一帧 |
| **无停顿** | 信息密度满格 | 关键节点留 300ms+ 呼吸 |
| **过度炫技** | 每页都有复杂动画 | 全篇只有一处 "120% 精致" |
| **魔法感** | 隐藏工作过程 | 展示 tweak、展示 bug 修复 |
| **缺少重量** | 元素飘在空中 | 落地有惯性（expoOut） |
| **线性运动** | 直线位移 | 弧线轨迹 + 自然曲线 |

### ✅ 必须遵循（12 项）

- [ ] 叙事节奏是 Slow-Fast-Boom-Stop
- [ ] 默认 easing 是 xpoOut（keySplines="0.19 1 0.22 1"）
- [ ] Toggle/按钮用了 overshoot（keySplines="0.34 1.56 0.64 1"）
- [ ] 列表有 30ms+ stagger
- [ ] 关键结果前有 0.5s 悬停
- [ ] 打字用 Chunk Reveal（词组为单位）
- [ ] 焦点切换加了 blur（不只是 opacity）
- [ ] Logo 是 Morph 形变
- [ ] 底色带色温（非纯黑白）
- [ ] 文字有衬线 + 无衬线层次
- [ ] 收尾是戛然而止
- [ ] 鼠标轨迹是弧线

---

## 五、60 秒自检清单

完成动画后，逐条确认：

### 叙事节奏
- [ ] 有 Slow 建立预期？
- [ ] 有 Fast 建立动量？
- [ ] 有 Boom 情感峰值？
- [ ] 有 Stop 决定感收尾？

### 运动曲线
- [ ] 默认用 expoOut？
- [ ] 按钮/卡片用 overshoot？
- [ ] 没有 linear 或 ease？

### 时间分配
- [ ] 列表项有 stagger？
- [ ] 关键结果前有悬停？
- [ ] 总时长 ≤ 3 秒/页？

### 视觉工艺
- [ ] 底色带色温？
- [ ] 文字有层次？
- [ ] 收尾是戛然而止？

### 无障碍
- [ ] 对比度 ≥ 4.5:1？
- [ ] 动画可暂停？
- [ ] 不影响阅读？

---

## 六、审美评级标准

完成动画后，问自己：**观众看完第一反应是什么？**

| 观众反应 | 评级 | 诊断 |
|---|---|---|
| "看起来挺流畅的" | good | 合格但无特色，你在做 PowerPoint |
| "这个动画真顺" | good+ | 技术对了，但没惊艳 |
| "这个东西看起来真的像**从桌面上浮起来的**" | great | 你触到了物理重量感 |
| "这不像是 AI 做的" | great+ | 你触到了 Anthropic 的门槛 |
| "我想**截图**发朋友圈" | great++ | 你做到了让观众主动传播 |

**great 和 good 的区别，不在于技术正确度，在于品味判断。**

---

## 七、与基础技能的集成

### 集成到 ppt-animate

本规范通过以下方式增强基础动画技能：

1. **增强默认模板**
   - 简洁商务风格：duration 0.4-0.6s，stagger 0.08s
   - 动态活力风格：duration 0.3-0.4s，stagger 0.06s
   - 优雅柔和风格：duration 0.6-0.8s，stagger 0.12s
   - 科技未来风格：duration 0.5s，stagger 0.10s

2. **智能动画选择**
   - 根据页面内容类型自动推荐动画组合
   - 根据产品性格配方（A/B/C）调整参数

3. **质量检查**
   - 运行时自动检测 16 条避坑清单
   - 输出动画配置报告

### 配置示例

`python
from ppt_animate import animate_pptx

# 使用平衡型配方（默认）
animate_pptx(
    'input.pptx',
    'output.pptx',
    style='balanced',  # 或 'lightweight', 'epic'
    method='xml'
)

# 自定义动画参数
animate_pptx(
    'input.pptx',
    'output.pptx',
    transitions=['fade', 'wipe', 'zoom'],
    entrance_timing={
        'title': {'duration': 0.8, 'easing': 'expoOut'},
        'list': {'duration': 0.4, 'stagger': 0.08},
        'chart': {'duration': 0.6, 'easing': 'spring'}
    },
    method='xml'
)
`

---

**版本**：v1.0（整合自 huashu-design animation-best-practices.md）
**关联文档**：
- D:\Workspace\PPT设计指南\animation-standards.md（详细参考）
- SKILL.md（基础动画技能）
