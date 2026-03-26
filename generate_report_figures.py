from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch


ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "report_figures"

FONT_CANDIDATES = [
    "PingFang SC",
    "Hiragino Sans GB",
    "Source Han Sans SC",
    "Noto Sans CJK SC",
    "Microsoft YaHei",
    "SimHei",
    "Arial Unicode MS",
]

PALETTE = {
    "bg": "#f7f1e5",
    "panel": "#fffaf0",
    "line": "#d6c3a5",
    "text": "#2f2517",
    "muted": "#6f5a3f",
    "accent": "#8b5e34",
    "accent_soft": "#f2dfbf",
    "green": "#dbe9dc",
    "blue": "#dce8f6",
    "gold": "#f4e7c8",
}


def get_fontproperties():
    for name in FONT_CANDIDATES:
        try:
            path = font_manager.findfont(name, fallback_to_default=False)
        except Exception:
            continue
        if path:
            return font_manager.FontProperties(fname=path)
    return font_manager.FontProperties()


FONT = get_fontproperties()


def setup_figure(width: float, height: float):
    fig, ax = plt.subplots(figsize=(width, height), dpi=200)
    fig.patch.set_facecolor(PALETTE["bg"])
    ax.set_facecolor(PALETTE["bg"])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")
    return fig, ax


def draw_title(ax, title: str, subtitle: str):
    ax.text(
        0.05,
        0.95,
        title,
        fontproperties=FONT,
        fontsize=20,
        fontweight="bold",
        color=PALETTE["text"],
        va="top",
    )
    ax.text(
        0.05,
        0.905,
        subtitle,
        fontproperties=FONT,
        fontsize=10.5,
        color=PALETTE["muted"],
        va="top",
    )


def draw_box(ax, x: float, y: float, w: float, h: float, title: str, body: str, fill: str):
    rect = FancyBboxPatch(
        (x, y),
        w,
        h,
        boxstyle="round,pad=0.012,rounding_size=0.025",
        linewidth=1.6,
        edgecolor=PALETTE["line"],
        facecolor=fill,
    )
    ax.add_patch(rect)

    if body.strip():
        ax.text(
            x + 0.02,
            y + h - 0.05,
            title,
            fontproperties=FONT,
            fontsize=12.5,
            fontweight="bold",
            color=PALETTE["text"],
            va="top",
        )
        ax.text(
            x + 0.02,
            y + h - 0.1,
            body,
            fontproperties=FONT,
            fontsize=9.8,
            color=PALETTE["text"],
            va="top",
            linespacing=1.5,
        )
    else:
        ax.text(
            x + w / 2,
            y + h / 2,
            title,
            fontproperties=FONT,
            fontsize=12.8,
            fontweight="bold",
            color=PALETTE["text"],
            ha="center",
            va="center",
        )


def draw_arrow(ax, start: tuple[float, float], end: tuple[float, float], label: str | None = None):
    arrow = FancyArrowPatch(
        start,
        end,
        arrowstyle="-|>",
        mutation_scale=16,
        linewidth=1.8,
        color=PALETTE["accent"],
        shrinkA=4,
        shrinkB=4,
    )
    ax.add_patch(arrow)

    if label:
        mid_x = (start[0] + end[0]) / 2
        mid_y = (start[1] + end[1]) / 2
        ax.text(
            mid_x,
            mid_y + 0.03,
            label,
            fontproperties=FONT,
            fontsize=9.2,
            color=PALETTE["muted"],
            ha="center",
            va="center",
        )


def save_figure(fig, stem: str):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    png_path = OUTPUT_DIR / f"{stem}.png"
    svg_path = OUTPUT_DIR / f"{stem}.svg"
    fig.savefig(png_path, bbox_inches="tight", facecolor=fig.get_facecolor())
    fig.savefig(svg_path, bbox_inches="tight", facecolor=fig.get_facecolor())
    plt.close(fig)
    return png_path, svg_path


def generate_figure_2_1():
    fig, ax = setup_figure(13, 3.8)
    draw_title(
        ax,
        "图2-1 文本梳理流程图",
        "对应论文第2.3节，展示文本梳理的核心处理步骤。",
    )

    boxes = [
        (0.05, 0.38, 0.125, 0.18, "原文分句", "", PALETTE["gold"]),
        (0.205, 0.38, 0.125, 0.18, "表格整理", "", PALETTE["panel"]),
        (0.36, 0.38, 0.125, 0.18, "人工提取", "", PALETTE["green"]),
        (0.515, 0.38, 0.125, 0.18, "机器分类", "", PALETTE["blue"]),
        (0.67, 0.38, 0.125, 0.18, "人工校对", "", PALETTE["green"]),
        (0.825, 0.38, 0.125, 0.18, "格式校核", "", PALETTE["gold"]),
    ]

    for box in boxes:
        draw_box(ax, *box)

    for i in range(len(boxes) - 1):
        x1 = boxes[i][0] + boxes[i][2]
        y1 = boxes[i][1] + boxes[i][3] / 2
        x2 = boxes[i + 1][0]
        y2 = boxes[i + 1][1] + boxes[i + 1][3] / 2
        draw_arrow(ax, (x1, y1), (x2, y2))
    return save_figure(fig, "figure_2_1_text_workflow")


def generate_figure_2_2():
    fig, ax = setup_figure(12.5, 8)
    draw_title(
        ax,
        "图2-2 网页实现逻辑图",
        "对应论文第2.5节，展示结构化数据如何经由脚本与前端页面转化为可交互的数字展示平台。",
    )

    draw_box(
        ax,
        0.08,
        0.68,
        0.22,
        0.16,
        "数据来源层",
        "原文分句表\n注释表\n阐释表\n评价关系表\n哲学词条表",
        PALETTE["gold"],
    )
    draw_box(
        ax,
        0.39,
        0.66,
        0.22,
        0.20,
        "数据生成层",
        "generate_web_data.py\ngenerate_criticism_data.py\ngenerate_philosophy_data.py\n\n统一索引与字段结构",
        PALETTE["green"],
    )
    draw_box(
        ax,
        0.70,
        0.68,
        0.22,
        0.16,
        "前端数据层",
        "data.js\ncriticism_data.js\nphilosophy_data.js",
        PALETTE["blue"],
    )

    draw_box(
        ax,
        0.08,
        0.34,
        0.35,
        0.20,
        "页面渲染层",
        "index.html\nreader.html\ncriticism.html\nphilosophy.html\n\n按 text_id、sentence_id、句段范围调用数据",
        PALETTE["panel"],
    )
    draw_box(
        ax,
        0.56,
        0.34,
        0.36,
        0.20,
        "交互逻辑层",
        "篇章切换\n分句点击弹窗\n跨句阐释高亮联动\n关键词搜索\n儒/释/道筛选",
        PALETTE["panel"],
    )

    draw_box(
        ax,
        0.22,
        0.08,
        0.56,
        0.16,
        "结果呈现层",
        "首页总览 ｜ 总阅读器 ｜ 后人对前人的评价 ｜ 哲学词典\n实现原文、单句注释、跨句阐释、评价关系与概念词条的综合展示",
        PALETTE["accent_soft"],
    )

    draw_arrow(ax, (0.30, 0.76), (0.39, 0.76), "清洗与转换")
    draw_arrow(ax, (0.61, 0.76), (0.70, 0.76), "输出脚本数据")
    draw_arrow(ax, (0.81, 0.68), (0.34, 0.54), "加载")
    draw_arrow(ax, (0.43, 0.44), (0.56, 0.44), "事件驱动")
    draw_arrow(ax, (0.38, 0.34), (0.42, 0.24))
    draw_arrow(ax, (0.74, 0.34), (0.58, 0.24))

    ax.text(
        0.08,
        0.60,
        "核心实现：以前端静态页面承载展示逻辑，以结构化数据脚本保证内容可复用、可检索、可联动。",
        fontproperties=FONT,
        fontsize=10.2,
        color=PALETTE["muted"],
    )
    return save_figure(fig, "figure_2_2_web_pipeline")


def generate_figure_3_1():
    fig, ax = setup_figure(10.5, 6.8)
    draw_title(
        ax,
        "图3-1 网站结构树形图",
        "对应论文第3.1节，展示平台的主要页面模块及其基本层级关系。",
    )

    root = (0.38, 0.78, 0.24, 0.12, "庄子注疏平台", "", PALETTE["accent_soft"])
    children = [
        (0.08, 0.50, 0.18, 0.12, "首页", "", PALETTE["gold"]),
        (0.30, 0.50, 0.18, 0.12, "总阅读器", "", PALETTE["blue"]),
        (0.52, 0.50, 0.18, 0.12, "评价关系", "", PALETTE["green"]),
        (0.74, 0.50, 0.18, 0.12, "哲学词典", "", PALETTE["panel"]),
    ]
    leaves = [
        (0.05, 0.22, 0.12, 0.1, "平台概览", "", PALETTE["panel"]),
        (0.18, 0.22, 0.12, 0.1, "篇章入口", "", PALETTE["panel"]),
        (0.31, 0.22, 0.16, 0.1, "原文分句", "", PALETTE["panel"]),
        (0.49, 0.22, 0.16, 0.1, "单句注释", "", PALETTE["panel"]),
        (0.67, 0.22, 0.16, 0.1, "关系展示", "", PALETTE["panel"]),
        (0.85 - 0.12, 0.22, 0.12, 0.1, "概念词条", "", PALETTE["panel"]),
    ]

    draw_box(ax, *root)
    for box in children:
        draw_box(ax, *box)
    for box in leaves:
        draw_box(ax, *box)

    root_center = (root[0] + root[2] / 2, root[1])
    for box in children:
        child_center = (box[0] + box[2] / 2, box[1] + box[3])
        draw_arrow(ax, root_center, child_center)

    child_leaf_map = {
        0: [0, 1],
        1: [2, 3],
        2: [4],
        3: [5],
    }
    for child_idx, leaf_indices in child_leaf_map.items():
        child = children[child_idx]
        child_center = (child[0] + child[2] / 2, child[1])
        for leaf_idx in leaf_indices:
            leaf = leaves[leaf_idx]
            leaf_center = (leaf[0] + leaf[2] / 2, leaf[1] + leaf[3])
            draw_arrow(ax, child_center, leaf_center)

    return save_figure(fig, "figure_3_1_site_tree")


def main():
    outputs = []
    outputs.extend(generate_figure_2_1())
    outputs.extend(generate_figure_2_2())
    outputs.extend(generate_figure_3_1())

    print("Generated files:")
    for path in outputs:
        print(path.relative_to(ROOT))


if __name__ == "__main__":
    main()
