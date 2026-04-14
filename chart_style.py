"""
UNSPED Grafik Stili — Merkezi Modül
Dashboard ve Word export tarafından import edilir.
"""
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.gridspec import GridSpec
from matplotlib.patches import Rectangle
import numpy as np
import io

# ── Renk Paleti ────────────────────────────────────────────────
PAL = {
    "bg":     "#FFFFFF",
    "panel":  "#F8FAFC",
    "k1":     "#0099CC",
    "k2":     "#7C3AED",
    "k3":     "#E8453C",
    "text":   "#1A2B3C",
    "muted":  "#64748B",
    "grid":   "#E2E8F0",
    "gold":   "#D97706",
    "up":     "#DC2626",
    "down":   "#16A34A",
    "border": "#CBD5E1",
}
K_COLORS = [PAL["k1"], PAL["k2"], PAL["k3"]]
K_LABELS  = ["Kapsam 1", "Kapsam 2", "Kapsam 3"]

RCPARAMS = {
    "figure.facecolor":   PAL["bg"],
    "axes.facecolor":     PAL["panel"],
    "axes.edgecolor":     PAL["border"],
    "axes.linewidth":     0.8,
    "axes.spines.top":    False,
    "axes.spines.right":  False,
    "axes.spines.left":   True,
    "axes.spines.bottom": True,
    "axes.grid":          True,
    "grid.color":         PAL["grid"],
    "grid.linestyle":     ":",
    "grid.linewidth":     0.9,
    "font.family":        "DejaVu Sans",
    "text.color":         PAL["text"],
    "axes.titlecolor":    PAL["text"],
    "axes.labelcolor":    PAL["muted"],
    "xtick.color":        PAL["muted"],
    "ytick.color":        PAL["muted"],
    "legend.facecolor":   PAL["bg"],
    "legend.edgecolor":   PAL["border"],
    "legend.labelcolor":  PAL["text"],
    "legend.fontsize":    9,
}

def apply_style():
    plt.rcParams.update(RCPARAMS)

def fmt_k(v):
    return f"{v:,.0f}" if v else "0"

# ══════════════════════════════════════════════════════════════════
# GRAFİK 1 — Grouped Bar + Donut
# ══════════════════════════════════════════════════════════════════
def make_chart1(years, k1v, k2v, k3v, totv, dpi=160):
    apply_style()
    fig = plt.figure(figsize=(14, 7), facecolor=PAL["bg"])
    gs  = GridSpec(1, 2, figure=fig, wspace=0.42,
                   left=0.07, right=0.97, top=0.82, bottom=0.13)

    fig.text(0.5, 0.96, "UNSPED  ·  Sera Gazı Emisyon Analizi",
             ha="center", fontsize=13, fontweight="bold", color=PAL["text"])
    fig.text(0.5, 0.90, " · ".join(years),
             ha="center", fontsize=9, color=PAL["muted"])
    fig.add_artist(plt.Line2D([0.07,0.93],[0.88,0.88],
                   transform=fig.transFigure, color=PAL["k1"],
                   linewidth=2, solid_capstyle="round"))

    x = np.arange(len(years)); w = 0.22
    ax1 = fig.add_subplot(gs[0])

    for i, (vals, col, lbl) in enumerate(zip([k1v,k2v,k3v], K_COLORS, K_LABELS)):
        pos = x + (i-1)*w
        ax1.bar(pos, vals, width=w, color=col, label=lbl,
                alpha=0.88, zorder=3, edgecolor=PAL["bg"], linewidth=1.2)
        ax1.bar(pos, vals, width=w*1.9, color=col, alpha=0.06, zorder=2)
        for p, h in zip(pos, vals):
            if h > 30:
                max_h = max(totv)
                if h > max_h * 0.12:
                    ax1.text(p, h*0.5, f"{h:,.0f}", ha="center", va="center",
                             fontsize=7.5, color="white", fontweight="bold",
                             rotation=90)
                else:
                    ax1.text(p, h+max_h*0.009, f"{h:,.0f}", ha="center",
                             va="bottom", fontsize=8, color=col, fontweight="bold")

    ax1.set_xticks(x); ax1.set_xticklabels(years, fontsize=10)
    ax1.set_ylabel("ton CO₂e", fontsize=9)
    ax1.set_title("Yıllık Kapsam Karşılaştırması", fontsize=12,
                  fontweight="bold", pad=10)
    ax1.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v,_: f"{v:,.0f}"))
    ax1.legend(loc="upper right", fontsize=9, framealpha=0.95)
    ax1.set_ylim(0, max(totv)*1.22)
    for i, tot in enumerate(totv):
        ax1.text(i, -max(totv)*0.07, f"{tot:,.0f}t",
                 ha="center", fontsize=8, color=PAL["gold"], fontweight="bold")

    # Donut
    ax2 = fig.add_subplot(gs[1])
    pie_vals = [k1v[-1], k2v[-1], k3v[-1]]
    wedges, _, autos = ax2.pie(
        pie_vals, colors=K_COLORS,
        autopct=lambda p: f"%{p:.1f}" if p > 5 else "",
        startangle=120, pctdistance=0.68,
        textprops={"fontsize":9},
        wedgeprops={"linewidth":3,"edgecolor":PAL["bg"],"width":0.55},
        explode=[0.04 if v/sum(pie_vals)<0.15 else 0 for v in pie_vals]
    )
    for a in autos:
        a.set_fontsize(10); a.set_color("white"); a.set_fontweight("bold")

    for wedge, lbl, val, col in zip(wedges, K_LABELS, pie_vals, K_COLORS):
        ang   = np.deg2rad((wedge.theta1+wedge.theta2)/2)
        cos_a = np.cos(ang); sin_a = np.sin(ang)
        ax2.annotate(f"{lbl}\n{val:,.0f}t",
            xy=(cos_a*0.78, sin_a*0.78),
            xytext=(cos_a*1.38, sin_a*1.38),
            ha="center", va="center", fontsize=9,
            color=col, fontweight="bold",
            arrowprops={"arrowstyle":"-","color":PAL["border"],"lw":0.9},
            bbox={"boxstyle":"round,pad=0.2","fc":PAL["bg"],
                  "ec":PAL["border"],"lw":0.6})

    ax2.text(0, 0.08, f"{sum(pie_vals):,.0f}", ha="center", va="center",
             fontsize=15, fontweight="bold", color=PAL["text"])
    ax2.text(0, -0.20, "ton CO₂e", ha="center", va="center",
             fontsize=9, color=PAL["muted"])
    ax2.text(0, 0.35, years[-1], ha="center", va="center",
             fontsize=11, fontweight="bold", color=PAL["gold"])
    ax2.set_title("Emisyon Dağılımı", fontsize=12, fontweight="bold", pad=10)
    ax2.set_xlim(-1.7, 1.7); ax2.set_ylim(-1.6, 1.6)
    return fig

# ══════════════════════════════════════════════════════════════════
# GRAFİK 2 — Stacked Bar + Trend Çizgisi
# ══════════════════════════════════════════════════════════════════
def make_chart2(years, k1v, k2v, k3v, totv, dpi=160):
    apply_style()
    fig2, ax3 = plt.subplots(figsize=(10, 6), facecolor=PAL["bg"])
    ax3.set_facecolor(PAL["panel"])
    ax3.spines["top"].set_visible(False)
    ax3.spines["right"].set_visible(False)
    ax3.spines["left"].set_color(PAL["border"])
    ax3.spines["bottom"].set_color(PAL["border"])

    x = np.arange(len(years))
    bottoms = np.zeros(len(years))
    for vals, col, lbl in zip([k1v,k2v,k3v], K_COLORS, K_LABELS):
        v = np.array(vals)
        ax3.bar(x, v, bottom=bottoms, color=col, label=lbl,
                alpha=0.88, zorder=3, edgecolor=PAL["bg"],
                linewidth=1.5, width=0.52)
        ax3.bar(x, v, bottom=bottoms, color=col, alpha=0.07,
                zorder=2, width=0.72)
        bottoms += v

    ax3.plot(x, totv, color=PAL["gold"], linewidth=2.5,
             marker="D", markersize=9, zorder=5, label="Toplam",
             markerfacecolor=PAL["bg"], markeredgecolor=PAL["gold"],
             markeredgewidth=2.2)

    for i in range(len(years)):
        ax3.text(i, totv[i]+max(totv)*0.025, f"{totv[i]:,.0f}",
                 ha="center", va="bottom", fontsize=9,
                 color=PAL["gold"], fontweight="bold")
        if i > 0:
            dp  = (totv[i]-totv[i-1])/totv[i-1]*100
            sym = "▲" if dp>0 else "▼"
            col = PAL["up"] if dp>0 else PAL["down"]
            ax3.text(i, max(totv[i],totv[i-1])+max(totv)*0.11,
                     f"{sym} %{abs(dp):.0f}",
                     ha="center", fontsize=13, fontweight="bold", color=col)

    ax3.set_xticks(x); ax3.set_xticklabels(years, fontsize=11)
    ax3.set_ylabel("ton CO₂e", fontsize=9)
    ax3.set_title("Yıllık Toplam Emisyon Trendi", fontsize=13,
                  fontweight="bold", pad=14)
    ax3.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v,_: f"{v:,.0f}"))
    ax3.set_ylim(0, max(totv)*1.38)
    ax3.legend(loc="upper right", ncol=2, fontsize=9, framealpha=0.95)
    fig2.text(0.5, 0.95, "Kapsamlara göre yığılmış  ·  ◆ Toplam eğrisi",
              ha="center", fontsize=9, color=PAL["muted"])
    fig2.tight_layout(rect=[0,0,1,0.93])
    return fig2

# ══════════════════════════════════════════════════════════════════
# GRAFİK 3 — KPI Kartları
# ══════════════════════════════════════════════════════════════════
def make_kpi(years, k1v, k2v, k3v, totv, dpi=160):
    apply_style()
    fig3 = plt.figure(figsize=(14, 3.2), facecolor=PAL["bg"])
    kpi_data = [
        ("Toplam Emisyon", totv[-1], totv[-2] if len(totv)>1 else 0, PAL["gold"]),
        ("Kapsam 1",       k1v[-1],  k1v[-2]  if len(k1v)>1  else 0, PAL["k1"]),
        ("Kapsam 2",       k2v[-1],  k2v[-2]  if len(k2v)>1  else 0, PAL["k2"]),
        ("Kapsam 3",       k3v[-1],  k3v[-2]  if len(k3v)>1  else 0, PAL["k3"]),
    ]
    for i, (lbl, curr, prev, col) in enumerate(kpi_data):
        ax = fig3.add_axes([0.02+i*0.248, 0.06, 0.225, 0.86])
        ax.set_facecolor(PAL["panel"])
        for sp in ax.spines.values():
            sp.set_color(PAL["border"]); sp.set_linewidth(0.8)
        ax.set_xticks([]); ax.set_yticks([])
        rect = Rectangle((0, 0.05), 0.045, 0.90,
                          transform=ax.transAxes, color=col,
                          zorder=5, clip_on=False)
        ax.add_patch(rect)
        dp      = (curr-prev)/prev*100 if prev else 0
        sym     = "▲" if dp > 0 else "▼"
        chg_col = PAL["up"] if dp > 0 else PAL["down"]
        ax.text(0.55, 0.68, f"{curr:,.0f}", ha="center", va="center",
                fontsize=17, fontweight="bold", color=PAL["text"],
                transform=ax.transAxes)
        ax.text(0.55, 0.47, "ton CO₂e", ha="center", va="center",
                fontsize=8, color=PAL["muted"], transform=ax.transAxes)
        ax.text(0.55, 0.28, lbl, ha="center", va="center",
                fontsize=10, fontweight="bold", color=col,
                transform=ax.transAxes)
        ax.text(0.55, 0.10,
                f"{sym} %{abs(dp):.0f}  ({years[-2] if len(years)>1 else '—'}→{years[-1]})",
                ha="center", va="center", fontsize=8.5,
                fontweight="bold", color=chg_col, transform=ax.transAxes)
    fig3.text(0.5, 0.99, f"KPI Özeti  ·  {years[-1]}",
              ha="center", va="top", fontsize=11,
              fontweight="bold", color=PAL["text"])
    return fig3

def fig_to_bytes(fig, dpi=160):
    """Figürü PNG bytes olarak döndür."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi,
                bbox_inches="tight", facecolor=PAL["bg"])
    plt.close(fig)
    buf.seek(0)
    return buf.read()

def fig_to_buf(fig, dpi=160):
    """Figürü BytesIO olarak döndür (Word için)."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi,
                bbox_inches="tight", facecolor=PAL["bg"])
    plt.close(fig)
    buf.seek(0)
    return buf
