# -*- coding: utf-8 -*-
"""生成《留量大脑》作品书配图。

图表统一使用黑体（SimHei），与正文宋体形成层级区分；
配色与文档一致（墨绿系），保证图文风格统一。
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Polygon
from matplotlib.font_manager import FontProperties

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.dirname(HERE)  # 项目根目录
FIGDIR = os.path.join(OUT, "figures")
os.makedirs(FIGDIR, exist_ok=True)

FONT_PATH = r"C:\Windows\Fonts\simhei.ttf"
if not os.path.exists(FONT_PATH):
    FONT_PATH = r"C:\Windows\Fonts\msyh.ttc"
fp = lambda s: FontProperties(fname=FONT_PATH, size=s)
plt.rcParams["axes.unicode_minus"] = False

ACCENT = "#1B4D3E"
ACCENT2 = "#2E7D64"
ACCENT3 = "#5FA98C"
FILL1 = "#E7F1EC"
FILL2 = "#D3E7DD"
GREY = "#595959"
LGREY = "#BFBFBF"
WARN = "#C0504D"


def save(fig, name):
    path = os.path.join(FIGDIR, name)
    fig.savefig(path, dpi=300, bbox_inches="tight", facecolor="white",
                pad_inches=0.12)
    plt.close(fig)
    print("saved", path)


# ================================================== 图 5-1 增速回落与成本抬升
def fig_cost():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(6.1, 3.0))
    fig.subplots_adjust(bottom=0.30, wspace=0.28)

    labels = ["2018", "2023", "2024"]
    vals = [600, 40.5, 15.0]
    bars = ax1.bar(labels, vals, color=[ACCENT3, ACCENT2, ACCENT], width=0.52)
    for b, v in zip(bars, vals):
        ax1.text(b.get_x() + b.get_width() / 2, v + 22, ("%g%%" % v),
                 ha="center", va="bottom", fontproperties=fp(8.5), color=ACCENT)
    ax1.set_title("市场增速快速回落（%）", fontproperties=fp(10), color=ACCENT, pad=8)
    ax1.set_ylim(0, 730)
    ax1.set_yticks([])
    for s in ("top", "right", "left"):
        ax1.spines[s].set_visible(False)
    ax1.spines["bottom"].set_color(LGREY)
    for lb in ax1.get_xticklabels():
        lb.set_fontproperties(fp(9))
        lb.set_color(GREY)

    labels2 = ["中小商家\n单客成本", "行业\n获客均值", "头部平台\n2022 口径"]
    vals2 = [450, 800, 1302]
    bars2 = ax2.bar(labels2, vals2, color=[ACCENT3, ACCENT2, ACCENT], width=0.52)
    for b, v, txt in zip(bars2, vals2, ["约 400–500 元", "约 800 元", "约 1302 元"]):
        ax2.text(b.get_x() + b.get_width() / 2, v + 40, txt,
                 ha="center", va="bottom", fontproperties=fp(8.2), color=ACCENT)
    ax2.set_title("获客成本持续抬升（元/人）", fontproperties=fp(10), color=ACCENT, pad=8)
    ax2.set_ylim(0, 1650)
    ax2.set_yticks([])
    for s in ("top", "right", "left"):
        ax2.spines[s].set_visible(False)
    ax2.spines["bottom"].set_color(LGREY)
    for lb in ax2.get_xticklabels():
        lb.set_fontproperties(fp(7.8))
        lb.set_color(GREY)

    fig.text(0.5, 0.035,
             "数据来源：公开行业研究报告与媒体报道，各口径不一致，仅示趋势方向（详见附录 C）",
             ha="center", va="bottom", fontproperties=fp(7.0), color=GREY)
    save(fig, "fig_cost.png")


# ================================================== 图 5-2 漏斗（漏底）
def fig_funnel():
    fig, ax = plt.subplots(figsize=(6.3, 3.6))
    ax.set_xlim(0, 13.4)
    ax.set_ylim(0, 11.0)
    ax.axis("off")

    CX = 5.2  # 漏斗中心

    def band(top, bottom, y0, y1, label, sub, fill, lw=1.1):
        p = Polygon([(CX - top / 2, y1), (CX + top / 2, y1),
                     (CX + bottom / 2, y0), (CX - bottom / 2, y0)],
                    closed=True, facecolor=fill, edgecolor=ACCENT2, lw=lw)
        ax.add_patch(p)
        if sub:
            ax.text(CX, (y0 + y1) / 2 + 0.20, label, ha="center", va="center",
                    fontproperties=fp(9.5), color=ACCENT)
            ax.text(CX, (y0 + y1) / 2 - 0.26, sub, ha="center", va="center",
                    fontproperties=fp(7.2), color=GREY)
        else:
            ax.text(CX, (y0 + y1) / 2, label, ha="center", va="center",
                    fontproperties=fp(9.0), color=ACCENT)

    # 引水端（宽）：公域触达 → 成交
    band(8.8, 8.0, 7.7, 9.2, "公域触达流量", "曝光 ／ 直播间 UV ／ 进店访客", FILL2)
    band(8.0, 7.2, 6.2, 7.7, "成交顾客", "完成一次购买", FILL1)

    # 漏水断口
    for sx, sy, ex in ((CX + 4.6, 6.2, CX + 2.1), (CX - 4.6, 6.2, CX - 2.1)):
        ax.add_patch(FancyArrowPatch((sx, sy), (ex, 5.55),
                                     arrowstyle="-|>", mutation_scale=10,
                                     color=WARN, lw=1.5, ls=(0, (4, 2))))
    ax.text(CX + 4.75, 5.95, "漏水点", ha="left", va="center",
            fontproperties=fp(10), color=WARN)
    ax.text(CX + 4.75, 5.30, "成交之后没有承接设计，\n访客“过而不留”", ha="left", va="top",
            fontproperties=fp(7.6), color=WARN)

    # 蓄水端（窄）：留量 → 复购
    band(4.0, 3.2, 4.5, 5.5, "可反复触达的留量", "加私域 ／ 会员 ／ 台账", FILL2)
    band(3.2, 2.5, 3.3, 4.5, "持续复购的客户", "老客贡献 ／ 转介绍", FILL1)

    ax.text(0.15, 10.75, "引水端投入密集",
            ha="left", va="top", fontproperties=fp(9.0), color=ACCENT2)
    ax.text(0.15, 10.10, "蓄水端几乎没有承接设计",
            ha="left", va="top", fontproperties=fp(9.0), color=WARN)
    save(fig, "fig_funnel.png")


# ================================================== 图 6-1 四阶沉淀模型
def fig_model():
    fig, ax = plt.subplots(figsize=(6.3, 3.5))
    ax.set_xlim(0, 13.6)
    ax.set_ylim(0, 8.6)
    ax.axis("off")

    stages = [
        ("一阶  引水", "获 取", "投流 · 内容 · 直播", "触达量 / 流量成本率"),
        ("二阶  蓄水", "留 存", "关注 · 加私域 · 会员", "沉淀率 / 加私域转化"),
        ("三阶  养水", "激 活", "分层触达 · 服务", "复购率 / 老客贡献率"),
        ("四阶  活水", "复 利", "数据回流 · 供应链协同", "私域贡献率 / 周转效率"),
    ]
    bw, bh, gap = 3.05, 2.5, 0.36
    x0, y = 0.15, 4.25
    for i, (t, en, acts, kpi) in enumerate(stages):
        x = x0 + i * (bw + gap)
        ax.add_patch(FancyBboxPatch((x, y), bw, bh,
                                    boxstyle="round,pad=0.05,rounding_size=0.12",
                                    facecolor=FILL1, edgecolor=ACCENT2, lw=1.3))
        ax.text(x + bw / 2, y + bh - 0.44, t, ha="center", va="center",
                fontproperties=fp(10.2), color=ACCENT)
        ax.text(x + bw / 2, y + bh - 0.98, en, ha="center", va="center",
                fontproperties=fp(7.4), color=ACCENT3)
        ax.text(x + bw / 2, y + bh - 1.58, acts, ha="center", va="center",
                fontproperties=fp(7.0), color=GREY)
        ax.text(x + bw / 2, y + 0.38, kpi, ha="center", va="center",
                fontproperties=fp(6.8), color=ACCENT2)
        if i < 3:
            ax.add_patch(FancyArrowPatch((x + bw + 0.03, y + bh / 2),
                                         (x + bw + gap - 0.03, y + bh / 2),
                                         arrowstyle="-|>", mutation_scale=10,
                                         color=ACCENT, lw=1.5))

    # 三个闸门（位于四个方框之间的间隙）
    for name, gi in (("入口闸门", 0), ("价值闸门", 1), ("数据闸门", 2)):
        gx = x0 + gi * (bw + gap) + bw + gap / 2
        ax.plot([gx, gx], [y + bh + 0.10, y + bh + 0.55], color=WARN, lw=1.1,
                ls=(0, (3, 2)))
        ax.text(gx, y + bh + 0.64, name, ha="center", va="bottom",
                fontproperties=fp(7.6), color=WARN)
    ax.text(x0, y + bh + 0.64, "闸门：", ha="left", va="bottom",
            fontproperties=fp(7.6), color=WARN)

    # 复利回路（走方框下方，避免压住正文）
    ax.add_patch(FancyArrowPatch((x0 + 4 * bw + 3 * gap - 0.4, y - 0.10),
                                 (x0 + 1.5, y - 0.10),
                                 connectionstyle="arc3,rad=-0.22",
                                 arrowstyle="-|>", mutation_scale=10,
                                 color=ACCENT3, lw=1.4, ls=(0, (5, 3))))
    ax.text(6.8, y - 2.05, "复利回路：关系与数据反哺上游，成本逐期下降",
            ha="center", va="center", fontproperties=fp(7.4), color=ACCENT2)

    ax.text(0.15, 8.45, "引水—蓄水—养水—活水：一条顺序不可颠倒的能力链",
            ha="left", va="top", fontproperties=fp(9.2), color=ACCENT)
    save(fig, "fig_model.png")

# ================================================== 图 8-1 平台五层架构
def fig_arch():
    fig, ax = plt.subplots(figsize=(6.3, 4.2))
    ax.set_xlim(0, 14.2)
    ax.set_ylim(0, 10.6)
    ax.axis("off")

    layers = [
        ("L5  应用服务层", "商户端小程序 · 治理驾驶舱 · 产业看板", ACCENT),
        ("L4  治理协同层", "聚合指标 · 风险预警 · 订单 / 产能 / 物流协同", ACCENT),
        ("L3  智能决策层", "沉淀诊断 · 策略生成 · 流失预警", ACCENT2),
        ("L2  沉淀度量层", "六项核心指标统一口径计算（沉淀率 / 复购率 …）", ACCENT2),
        ("L1  数据采集层", "授权范围内归集经营与互动数据（结构化 + 脱敏）", ACCENT2),
    ]
    x, w = 3.05, 7.6
    h, gap = 1.32, 0.22
    top = 8.9
    for i, (name, duty, col) in enumerate(layers):
        y = top - i * (h + gap) - h
        ax.add_patch(FancyBboxPatch((x, y), w, h,
                                    boxstyle="round,pad=0.05,rounding_size=0.12",
                                    facecolor=col, edgecolor="none"))
        ax.text(x + 0.30, y + h - 0.42, name, ha="left", va="center",
                fontproperties=fp(9.4), color="white")
        ax.text(x + 0.30, y + 0.42, duty, ha="left", va="center",
                fontproperties=fp(7.1), color="#DCEDE4")

    # 数据边界：L3 与 L4 之间
    yb = top - 3 * (h + gap) + gap / 2
    ax.plot([1.55, 12.05], [yb, yb], color=WARN, lw=1.6, ls=(0, (4, 3)))
    ax.text(x + w / 2, yb, "数 据 边 界", ha="center", va="center",
            fontproperties=fp(8.0), color=WARN,
            bbox=dict(facecolor="white", edgecolor="none", pad=2.5))

    # 左侧：商户侧（覆盖 L3—L1）
    y_low = (top - 3 * (h + gap) + (top - 4 * (h + gap) - h)) / 2
    ax.text(1.45, y_low, "商户侧（数据不出商户）", ha="center", va="center",
            fontproperties=fp(8.2), color=ACCENT2, rotation=90)
    # 右侧：协同与服务层（覆盖 L5—L4）
    y_high = ((top - h) + (top - 2 * (h + gap))) / 2
    ax.text(13.15, y_high, "协同与服务层（只见聚合数据）", ha="center", va="center",
            fontproperties=fp(8.2), color=WARN, rotation=-90)

    ax.text(0.15, 10.45,
            "留量大脑五层架构：L1—L3 在商户侧完成，对外只上送聚合或脱敏结果",
            ha="left", va="top", fontproperties=fp(8.8), color=ACCENT)
    save(fig, "fig_arch.png")


if __name__ == "__main__":
    fig_cost()
    fig_funnel()
    fig_model()
    fig_arch()
    print("ALL FIGURES DONE")
