# -*- coding: utf-8 -*-
"""
配色方案库 - 8套力量感配色方案
来源：Obsidian 知识库 YI HENG 配色积累
"""
from pptx.dml.color import RGBColor

COLOR_SCHEMES = {
    '冷银灰力量': {
        'name': '冷银灰力量',
        'colors': {
            '背景': RGBColor(0x18, 0x16, 0x17),      # 曜石黑 #181617
            '卡片': RGBColor(0x1F, 0x1D, 0x1E),      # 深黑底
            '主色': RGBColor(0xB5, 0xB8, 0xBB),      # 冷银灰 #B5B8BB
            '强调': RGBColor(0xC5, 0x1E, 0x2A),      # 警戒红 #C51E2A
            '文字': RGBColor(0xE0, 0xE0, 0xE0),      # 浅白
            '辅助': RGBColor(0x88, 0x88, 0x88),      # 中灰
        },
        'tags': '理性、金属感、秩序感',
        'description': '曜石黑底 + 冷银灰主色 + 警戒红点缀，适合科技、数据类汇报',
    },
    '深海骨白': {
        'name': '深海骨白',
        'colors': {
            '背景': RGBColor(0xE7, 0xE1, 0xD3),      # 骨白 #E7E1D3
            '卡片': RGBColor(0xFF, 0xFE, 0xFA),      # 纯白
            '主色': RGBColor(0x12, 0x23, 0x3A),      # 深海军蓝 #12233A
            '强调': RGBColor(0xE8, 0x53, 0x27),      # 火焰橙 #E85327
            '文字': RGBColor(0x2A, 0x2A, 0x2A),      # 深灰
            '辅助': RGBColor(0x66, 0x66, 0x66),      # 中灰
        },
        'tags': '自然、沉静、平衡感',
        'description': '骨白底 + 深海军蓝主色 + 火焰橙点缀，适合人文、自然主题',
    },
    '帝王紫电光': {
        'name': '帝王紫电光',
        'colors': {
            '背景': RGBColor(0x14, 0x15, 0x14),      # 墨黑 #141514
            '卡片': RGBColor(0x1A, 0x1B, 0x1A),      # 深黑底
            '主色': RGBColor(0x4C, 0x2A, 0x71),      # 帝王紫 #4C2A71
            '强调': RGBColor(0xAD, 0xCF, 0x37),      # 酸性绿 #ADCF37
            '文字': RGBColor(0xE8, 0xE8, 0xE8),      # 浅白
            '辅助': RGBColor(0x99, 0x88, 0xBB),      # 紫灰
        },
        'tags': '权威、深邃、电子感',
        'description': '墨黑底 + 帝王紫主色 + 酸性绿点缀，适合创意、数字主题',
    },
    '铁锈工业': {
        'name': '铁锈工业',
        'colors': {
            '背景': RGBColor(0x2C, 0x2C, 0x2B),      # 炭灰 #343635 变体
            '卡片': RGBColor(0x3A, 0x3A, 0x39),      # 深炭
            '主色': RGBColor(0x9D, 0x39, 0x28),      # 铁锈红 #9D3928
            '强调': RGBColor(0xCE, 0xCB, 0xC1),      # 混凝土白 #CECBC1
            '文字': RGBColor(0xF0, 0xEE, 0xE8),      # 米白
            '辅助': RGBColor(0xAA, 0x99, 0x88),      # 暖灰
        },
        'tags': '朴素、结构感、生命力',
        'description': '炭灰底 + 铁锈红主色 + 混凝土白点缀，适合工程、建筑主题',
    },
    '森林黄铜': {
        'name': '森林黄铜',
        'colors': {
            '背景': RGBColor(0x2A, 0x22, 0x1D),      # 深褐黑 #2A221D
            '卡片': RGBColor(0x32, 0x2A, 0x24),      # 深褐
            '主色': RGBColor(0x1F, 0x4A, 0x39),      # 森林绿 #1F4A39
            '强调': RGBColor(0xB1, 0x87, 0x39),      # 黄铜金 #B18739
            '文字': RGBColor(0xE8, 0xE0, 0xD0),      # 米白
            '辅助': RGBColor(0x88, 0x77, 0x55),      # 褐灰
        },
        'tags': '沉静、稳固、岁月感',
        'description': '深褐黑底 + 森林绿主色 + 黄铜金点缀，适合自然、历史主题',
    },
    '钴蓝猩红': {
        'name': '钴蓝猩红',
        'colors': {
            '背景': RGBColor(0xEF, 0xE8, 0xD6),      # 纸本白 #EFE8D6
            '卡片': RGBColor(0xFF, 0xFF, 0xFE),      # 纯白
            '主色': RGBColor(0x1C, 0x4A, 0x9C),      # 钴蓝 #1C4A9C
            '强调': RGBColor(0xDA, 0x2E, 0x2B),      # 猩红 #DA2E2B
            '文字': RGBColor(0x1A, 0x1A, 0x1A),      # 深黑
            '辅助': RGBColor(0x55, 0x55, 0x55),      # 中灰
        },
        'tags': '质朴、人文、行动感',
        'description': '纸本白底 + 钴蓝主色 + 猩红点缀，适合学术、人文主题',
    },
    '勃艮第暗金': {
        'name': '勃艮第暗金',
        'colors': {
            '背景': RGBColor(0x2A, 0x24, 0x22),      # 深灰黑
            '卡片': RGBColor(0x32, 0x2C, 0x2A),      # 深褐灰
            '主色': RGBColor(0x65, 0x1F, 0x2B),      # 勃艮第红 #651F2B
            '强调': RGBColor(0x96, 0x71, 0x30),      # 暗金 #967130
            '文字': RGBColor(0xE8, 0xE0, 0xD0),      # 米白
            '辅助': RGBColor(0x69, 0x66, 0x60),      # 岩石灰 #696660
        },
        'tags': '尊贵、历史感、权力感',
        'description': '深灰底 + 勃艮第红主色 + 暗金点缀，适合高端、历史主题',
    },
    '石墨电光青': {
        'name': '石墨电光青',
        'colors': {
            '背景': RGBColor(0x1F, 0x23, 0x26),      # 石墨黑 #1F2326
            '卡片': RGBColor(0x26, 0x2A, 0x2D),      # 深石墨
            '主色': RGBColor(0x1C, 0x23, 0x4B),      # 深靛蓝 #1C234B
            '强调': RGBColor(0x26, 0xCD, 0xCB),      # 电光青 #26CDCB
            '文字': RGBColor(0xE0, 0xE8, 0xEC),      # 冷白
            '辅助': RGBColor(0x55, 0x66, 0x77),      # 蓝灰
        },
        'tags': '硬朗、精密、科技',
        'description': '石墨黑底 + 深靛蓝主色 + 电光青点缀，适合科技、数据主题',
    },
}

# 默认配色
DEFAULT_SCHEME = COLOR_SCHEMES['冷银灰力量']


def get_color_scheme_names():
    """获取所有配色方案名称列表"""
    return list(COLOR_SCHEMES.keys())


def get_color_scheme(name):
    """获取指定配色方案"""
    if name in COLOR_SCHEMES:
        return COLOR_SCHEMES[name]
    if name == '默认':
        return DEFAULT_SCHEME
    return None


def apply_scheme(prs, scheme_name):
    """应用配色方案到演示文稿（背景 + 样式）"""
    scheme = get_color_scheme(scheme_name)
    if scheme is None:
        scheme = DEFAULT_SCHEME
    colors = scheme['colors']
    # 遍历所有幻灯片，修改背景
    for slide in prs.slides:
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = colors['背景']
    return scheme
