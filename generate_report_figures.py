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
    fig, ax = setup_figure(14, 5.8)
    draw_title(
        ax,
        "图2-1 文本梳理流程图",
        "对应论文第2.3节，展示《庄子》原文与注疏材料由初步整理到结构化数据形成的处理链条。",
    )

    boxes = [
        (0.05, 0.34, 0.14, 0.34, "原文分句", "以通行整理本为基础，\n将原文切分为可对应的\n分句单元。", PALETTE["gold"]),
        (0.22, 0.34, 0.14, 0.34, "表格建索引", "建立篇名、句序、原文等字段，\n形成逐句录入与比对的\n基础框架。", PALETTE["panel"]),
        (0.39, 0.34, 0.14, 0.34, "人工提取注疏", "从多版本注疏中逐句提取\n对应内容，汇总到统一表格。", PALETTE["green"]),
        (0.56, 0.34, 0.14, 0.34, "机器分类", "依据语言特征初步区分\n“训诂性注释”与“阐释性注释”。", PALETTE["blue"]),
        (0.73, 0.34, 0.14, 0.34, "人工校对", "结合上下文复核分类边界，\n修正误判与缺漏。", PALETTE["green"]),
        (0.90 - 0.14, 0.34, 0.14, 0.34, "数据格式化", "统一字段、来源与标签，\n完成最终校核，形成结构化数据。", PALETTE["gold"]),
    ]

    for box in boxes:
        draw_box(ax, *box)

    for i in range(len(boxes) - 1):
        x1 = boxes[i][0] + boxes[i][2]
        y1 = boxes[i][1] + boxes[i][3] / 2
        x2 = boxes[i + 1][0]
        y2 = boxes[i + 1][1] + boxes[i + 1][3] / 2
        draw_arrow(ax, (x1, y1), (x2, y2))

    ax.text(
        0.05,
        0.16,
        "输出结果：形成可用于网页渲染、文本检索与结果分析的统一数据表。",
        fontproperties=FONT,
        fontsize=11,
        color=PALETTE["text"],
    )
    ax.text(
        0.05,
        0.10,
        "说明：机器步骤承担初筛与提效功能，最终结果以人工校对后的数据为准。",
        fontproperties=FONT,
        fontsize=10,
        color=PALETTE["muted"],
    )
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


def main():
    outputs = []
    outputs.extend(generate_figure_2_1())
    outputs.extend(generate_figure_2_2())

    print("Generated files:")
    for path in outputs:
        print(path.relative_to(ROOT))


if __name__ == "__main__":
    main()
