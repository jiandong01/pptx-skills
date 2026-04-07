"""
标准布局分类体系

定义了 5 大结构性布局 + 7 种内容性布局的标准分类，
包括别名映射、关键词匹配规则等。
"""

from dataclasses import dataclass
from typing import List, Optional


@dataclass
class LayoutStandard:
    """标准布局定义"""
    name: str  # 标准名称
    aliases: List[str]  # 别名列表
    keywords: List[str]  # 模板名称关键词
    priority: int  # 自动匹配优先级（越高越优先）
    category: str  # 分类：structural 或 content
    description: str  # 描述


# 标准布局定义
STANDARD_LAYOUTS = {
    # ========== 结构性布局 ==========
    'cover': LayoutStandard(
        name='cover',
        aliases=['title-slide', '封面'],
        keywords=['title slide', '标题幻灯片', '封面'],
        priority=100,
        category='structural',
        description='封面页 - 演示文稿首页，展示标题、副标题'
    ),

    'toc': LayoutStandard(
        name='toc',
        aliases=['contents', 'agenda', '目录'],
        keywords=['目录', 'contents', 'agenda', 'toc'],
        priority=90,
        category='structural',
        description='目录页 - 展示演示文稿的章节结构'
    ),

    'section': LayoutStandard(
        name='section',
        aliases=['chapter', 'divider', '章节'],
        keywords=['section', '章节', '节标题', 'chapter'],
        priority=90,
        category='structural',
        description='章节页 - 标记新章节的开始'
    ),

    'summary': LayoutStandard(
        name='summary',
        aliases=['conclusion', 'ending', '总结', '致谢'],
        keywords=['summary', 'conclusion', '总结', '致谢'],
        priority=80,
        category='structural',
        description='总结页 - 演示文稿结尾，总结要点或致谢'
    ),

    # ========== 内容性布局 ==========
    'standard': LayoutStandard(
        name='standard',
        aliases=['title-content', 'default', '标准', '默认'],
        keywords=['title and content', '标题和内容', 'title & content'],
        priority=50,  # 默认布局，优先级较低
        category='content',
        description='标准布局 - 标题+正文，最常用'
    ),

    'two-column': LayoutStandard(
        name='two-column',
        aliases=['two-content', 'left-right', '双栏', '左右'],
        keywords=['two content', '双栏', '两栏', 'two-content'],
        priority=70,
        category='content',
        description='双栏布局 - 左右对比、并列展示'
    ),

    'image': LayoutStandard(
        name='image',
        aliases=['picture', 'visual', '图片'],
        keywords=['picture', 'image', '图片', 'photo'],
        priority=60,
        category='content',
        description='图片布局 - 展示图片、截图、示意图'
    ),

    'chart': LayoutStandard(
        name='chart',
        aliases=['graph', 'data', '图表'],
        keywords=['chart', 'graph', '图表'],
        priority=80,  # 图表优先级高
        category='content',
        description='图表布局 - 展示图表、数据可视化'
    ),

    'table': LayoutStandard(
        name='table',
        aliases=['comparison', '表格'],
        keywords=['table', '表格'],
        priority=60,
        category='content',
        description='表格布局 - 展示表格数据、对比信息'
    ),

    'mixed': LayoutStandard(
        name='mixed',
        aliases=['hybrid', 'text-image', '混合'],
        keywords=[],  # 通常复用 two-column
        priority=60,
        category='content',
        description='混合布局 - 文字+图片/图表混排'
    ),

    'title-only': LayoutStandard(
        name='title-only',
        aliases=['blank', 'free', '仅标题', '自由'],
        keywords=['title only', '仅标题'],
        priority=40,
        category='content',
        description='仅标题布局 - 自由布局，完全自定义'
    ),
}


def get_layout_standard(name: str) -> Optional[LayoutStandard]:
    """根据名称或别名获取标准布局定义"""
    # 直接匹配标准名称
    if name in STANDARD_LAYOUTS:
        return STANDARD_LAYOUTS[name]

    # 匹配别名
    name_lower = name.lower()
    for std_name, layout_std in STANDARD_LAYOUTS.items():
        if name_lower in [alias.lower() for alias in layout_std.aliases]:
            return layout_std

    return None


def find_standard_by_keywords(layout_name: str) -> Optional[str]:
    """根据模板布局名称的关键词，找到对应的标准布局名称"""
    layout_name_lower = layout_name.lower()

    # 按优先级排序
    sorted_layouts = sorted(
        STANDARD_LAYOUTS.items(),
        key=lambda x: x[1].priority,
        reverse=True
    )

    for std_name, layout_std in sorted_layouts:
        for keyword in layout_std.keywords:
            if keyword.lower() in layout_name_lower:
                return std_name

    return None


def get_all_structural_layouts() -> List[str]:
    """获取所有结构性布局名称"""
    return [name for name, std in STANDARD_LAYOUTS.items() if std.category == 'structural']


def get_all_content_layouts() -> List[str]:
    """获取所有内容性布局名称"""
    return [name for name, std in STANDARD_LAYOUTS.items() if std.category == 'content']


def print_layout_standards():
    """打印所有标准布局（用于调试）"""
    print("\n标准布局分类体系")
    print("=" * 60)

    print("\n结构性布局 (Structural Layouts):")
    for name in get_all_structural_layouts():
        std = STANDARD_LAYOUTS[name]
        print(f"  {name:15} - {std.description}")
        print(f"                  别名: {', '.join(std.aliases)}")

    print("\n内容性布局 (Content Layouts):")
    for name in get_all_content_layouts():
        std = STANDARD_LAYOUTS[name]
        print(f"  {name:15} - {std.description}")
        print(f"                  别名: {', '.join(std.aliases)}")

    print("\n" + "=" * 60)


if __name__ == '__main__':
    print_layout_standards()
