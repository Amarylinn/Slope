"""
TolSlope v2.0 - Analisis Stabilitas Lereng Timbunan Jalan Tol
Metode Keseimbangan Batas: Fellenius (1927) & Bishop Disederhanakan (1955)
Tema Tampilan Cerah (Bright Professional Light Theme) + Rumus LaTeX + Asal Data + Ekspor Max Excel & CAD
Jalankan lokal : streamlit run app.py
Deploy         : Upload app.py, requirements.txt, dokumentasi_tolslope_v2.txt ke GitHub
"""
import io, json
from datetime import datetime
import ezdxf
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.properties import PageSetupProperties

st.set_page_config(page_title="TolSlope v2.0 — Stabilitas Lereng Tol", page_icon="🏗️", layout="wide", initial_sidebar_state="expanded")

# ============================================================================
# CSS TAMPILAN CERAH (BRIGHT & CLEAN PROFESSIONAL THEME)
# ============================================================================
st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');
[data-testid="stAppViewContainer"]{background:linear-gradient(135deg,#f8fafc 0%,#f1f5f9 50%,#eef2f6 100%);color:#1e293b;font-family:'Plus Jakarta Sans',sans-serif}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#ffffff 0%,#f8fafc 100%);border-right:1px solid #e2e8f0}
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3{color:#1e3a8a!important}
.hero-title{font-size:2.35rem;font-weight:800;background:linear-gradient(90deg,#1d4ed8,#0284c7,#0d9488);-webkit-background-clip:text;-webkit-text-fill-color:transparent;text-align:center;margin-bottom:0;letter-spacing:-0.5px}
.hero-sub{text-align:center;color:#475569;font-size:0.95rem;margin-top:0.3rem;margin-bottom:1.3rem;font-weight:500}
.info-card{background:#ffffff;border:1px solid #e2e8f0;border-radius:14px;padding:1.2rem 1.4rem;margin-bottom:1rem;box-shadow:0 4px 16px rgba(0,0,0,0.03)}
.info-card h3{color:#1e3a8a;font-size:1.1rem;margin-bottom:0.5rem;font-weight:700}
.metric-row{display:flex;gap:0.8rem;flex-wrap:wrap;margin:1rem 0}
.metric-card{flex:1;min-width:170px;background:#ffffff;border:1px solid #e2e8f0;border-radius:12px;padding:1rem;text-align:center;box-shadow:0 3px 12px rgba(0,0,0,0.03);transition:transform 0.15s,box-shadow 0.15s}
.metric-card:hover{transform:translateY(-2px);box-shadow:0 6px 18px rgba(0,0,0,0.06)}
.metric-label{color:#64748b;font-size:0.75rem;text-transform:uppercase;letter-spacing:1px;font-weight:700}
.metric-val{font-size:1.95rem;font-weight:800;margin:0.2rem 0}
.safe{color:#059669}.danger{color:#dc2626}.warn{color:#d97706}.neutral{color:#0284c7}
.status-badge{display:inline-block;padding:0.45rem 1.4rem;border-radius:30px;font-weight:700;font-size:0.95rem;text-align:center;letter-spacing:0.5px}
.badge-safe{background:linear-gradient(135deg,#059669,#10b981);color:#fff;box-shadow:0 4px 14px rgba(16,185,129,0.35)}
.badge-warn{background:linear-gradient(135deg,#d97706,#f59e0b);color:#fff;box-shadow:0 4px 14px rgba(245,158,11,0.35)}
.badge-danger{background:linear-gradient(135deg,#dc2626,#ef4444);color:#fff;box-shadow:0 4px 14px rgba(239,68,68,0.35)}
.data-origin{display:inline-block;padding:0.18rem 0.65rem;border-radius:12px;font-size:0.7rem;font-weight:700;margin-top:0.3rem}
.origin-calc{background:#fef3c7;color:#92400e;border:1px solid #fde68a}
.origin-user{background:#dcfce7;color:#166534;border:1px solid #bbf7d0}
.origin-ref{background:#ede9fe;color:#5b21b6;border:1px solid #ddd6fe}
.formula-box{background:#f0f9ff;border-left:4px solid #0284c7;padding:0.9rem 1.2rem;border-radius:0 10px 10px 0;margin:0.6rem 0;font-size:0.88rem;line-height:1.6;color:#0f172a}
.stTabs [data-baseweb="tab-list"]{gap:0.4rem;background:#e2e8f0;border-radius:10px;padding:0.3rem}
.stTabs [data-baseweb="tab"]{border-radius:8px;font-weight:600;color:#475569}
.stTabs [aria-selected="true"]{background:#ffffff!important;color:#1d4ed8!important;box-shadow:0 2px 8px rgba(0,0,0,0.06)}
.stDownloadButton>button{background:linear-gradient(135deg,#1d4ed8,#0284c7)!important;color:#fff!important;border:none!important;border-radius:8px!important;font-weight:600!important;box-shadow:0 3px 10px rgba(2,132,199,0.25)}
</style>""", unsafe_allow_html=True)

def fmt(x, d=2): return f"{x:.{d}f}".replace(".", ",")

SOILS = {
    "Lunak": dict(t=6.0, g=15.5, c=20.0, ph=0.0, col="#e4d5b7", desc="Lempung lunak tak terdrainase (cu, φ=0)"),
    "Lempung": dict(t=5.0, g=17.0, c=50.0, ph=0.0, col="#c4aa82", desc="Lempung sedang-kaku"),
    "Keras": dict(t=5.0, g=19.0, c=25.0, ph=28.0, col="#967d5e", desc="Tanah keras/pasir padat efektif (c', φ')"),
}

def prep(p, names=None):
    p = dict(p)
    act = [i for i in (1, 2, 3) if p[f"t{i}"] > 0]
    p["Lt"] = np.array([p[f"t{i}"] for i in act], float)
    p["Lg"] = np.array([p[f"g{i}"] for i in act], float)
    p["Lc"] = np.array([p[f"c{i}"] for i in act], float)
    p["Lph"] = np.array([p[f"ph{i}"] for i in act], float)
    p["Ln"] = [(names[i - 1] if names else f"Lapisan {i}") for i in act]
    p["Lz"] = np.cumsum(p["Lt"])
    p["D"] = float(p["Lz"][-1])
    return p

def surface_y(x, p):
    return np.interp(x, [-1000.0, p["B"], p["B"] + p["m"] * p["H"], 1000.0], [p["H"], p["H"], 0.0, 0.0])

def circle_fs(p, xc, yc, R, nsl=50, detail=False):
    xs = np.linspace(xc - R, xc + R, 401)
    yarc = yc - np.sqrt(np.maximum(R * R - (xs - xc) ** 2, 0.0))
    d = surface_y(xs, p) - yarc
    idx = np.nonzero(d > 0)[0]
    if len(idx) < 2 or idx[0] == 0 or idx[-1] == len(xs) - 1 or not np.all(d[idx[0]:idx[-1] + 1] > 0): return None
    xa = xs[idx[0] - 1] + (0 - d[idx[0] - 1]) * (xs[idx[0]] - xs[idx[0] - 1]) / (d[idx[0]] - d[idx[0] - 1])
    xb = xs[idx[-1]] + (0 - d[idx[-1]]) * (xs[idx[-1] + 1] - xs[idx[-1]]) / (d[idx[-1] + 1] - d[idx[-1]])
    if xb <= p["B"] + 1e-6 or xb - xa < 1.0: return None
    edges = np.linspace(xa, xb, nsl + 1)
    xm = (edges[:-1] + edges[1:]) / 2
    b = (xb - xa) / nsl
    yb = yc - np.sqrt(R * R - (xm - xc) ** 2)
    ys = surface_y(xm, p)
    hE = np.maximum(0.0, ys - np.maximum(yb, 0.0))
    depth = np.maximum(0.0, -yb)
    top = p["Lz"] - p["Lt"]
    wF = np.sum(p["Lg"] * np.clip(depth[:, None] - top, 0.0, p["Lt"]), axis=1)
    Wsoil = b * (p["gE"] * hE + wF)
    W = Wsoil + p["q"] * b * (xm <= p["B"])
    sinA = (xc - xm) / R
    cosA = np.sqrt(np.maximum(1 - sinA ** 2, 1e-6))
    inE = yb >= 0
    idx_l = np.minimum(np.searchsorted(p["Lz"], depth, side="left"), len(p["Lz"]) - 1)
    c = np.where(inE, p["cE"], p["Lc"][idx_l])
    tphi = np.tan(np.radians(np.where(inE, p["phE"], p["Lph"][idx_l])))
    drive = float(np.sum(W * sinA))
    if drive <= 1e-9: return None
    ub = p["ru"] * Wsoil
    fel = float(np.sum(c * b / cosA + (W * cosA - ub / cosA) * tphi) / drive)
    if not np.isfinite(fel): return None
    fs = max(fel, 0.1)
    for _ in range(60):
        mA = cosA + sinA * tphi / fs
        if np.any(mA <= 0.05): return None
        new = float(np.sum((c * b + (W - ub) * tphi) / mA) / drive)
        if not np.isfinite(new) or new <= 0: return None
        if abs(new - fs) < 1e-5: fs = new; break
        fs = new
    out = dict(fs=fs, fel=fel, xa=xa, xb=xb, ya=float(surface_y(xa, p)), yb_=float(surface_y(xb, p)),
               drive=drive, resist_bishop=drive * fs, resist_fel=drive * fel, total_w=float(np.sum(W)))
    if detail:
        out["slices"] = pd.DataFrame({
            "x tengah (m)": xm, "lebar b (m)": b, "berat W (kN/m)": W,
            "α (°)": np.degrees(np.arcsin(sinA)), "c (kPa)": c, "φ (°)": np.degrees(np.arctan(tphi)),
            "W sin α": W * sinA, "c·b (kN/m)": c * b, "(W-ub)tanφ": (W - ub) * tphi
        }).round(3)
    return out

def search(p, nx=15, ny=11, nr=20, refine=True):
    H, m, B, D = p["H"], p["m"], p["B"], p["D"]
    toe = B + m * H
    xcs, ycs = np.linspace(B - 1.5 * H, toe + 0.5 * H, nx), np.linspace(0.9 * H, 2.5 * H + 0.8 * D, ny)
    grid = np.full((ny, nx), np.nan)
    best = None
    def scan(xc, yc, n):
        res = None
        for ybot in np.linspace(-D, 0.8 * H, n):
            R = yc - ybot
            if R < 1.0: continue
            r = circle_fs(p, xc, yc, R)
            if r is not None and (res is None or r["fs"] < res[0]): res = (r["fs"], R)
        return res
    for j, yc in enumerate(ycs):
        for i, xc in enumerate(xcs):
            r = scan(xc, yc, nr)
            if r is not None:
                grid[j, i] = r[0]
                if best is None or r[0] < best[0]: best = (r[0], xc, yc, r[1])
    if best is None: return dict(best=None, xcs=xcs, ycs=ycs, grid=grid)
    if refine:
        dx, dy = xcs[1] - xcs[0], ycs[1] - ycs[0]
        for xc in np.linspace(best[1] - dx, best[1] + dx, 7):
            for yc in np.linspace(best[2] - dy, best[2] + dy, 7):
                r = scan(xc, yc, 36)
                if r is not None and r[0] < best[0]: best = (r[0], xc, yc, r[1])
    return dict(best=best, xcs=xcs, ycs=ycs, grid=grid)

PKEYS = ["H", "B", "m", "gE", "cE", "phE", "q", "ru"] + [f"{k}{i}" for i in (1, 2, 3) for k in ("t", "g", "c", "ph")]

@st.cache_data(show_spinner="🔄 Mencari bidang longsor kritis...")
def cached_search(pt, nx, ny, nr, refine):
    return search(prep(dict(zip(PKEYS, pt))), nx, ny, nr, refine)

def arc_points(xc, yc, R, xa, xb, n=120):
    xs = np.linspace(xa, xb, n)
    return xs, yc - np.sqrt(np.maximum(R * R - (xs - xc) ** 2, 0.0))

def plot_result(p, res, target):
    fs, xc, yc, R = res["best"]
    r = circle_fs(p, xc, yc, R)
    fig, (ax, ax2) = plt.subplots(1, 2, figsize=(15, 5.6), gridspec_kw={"width_ratios": [1.55, 1]})
    fig.patch.set_facecolor('#ffffff')
    for a in (ax, ax2):
        a.set_facecolor('#f8fafc'); a.tick_params(colors='#334155'); a.xaxis.label.set_color('#1e293b')
        a.yaxis.label.set_color('#1e293b'); a.title.set_color('#1e3a8a')
        for s in a.spines.values(): s.set_color('#cbd5e1')
    toe = p["B"] + p["m"] * p["H"]
    xmin, xmax = min(r["xa"], p["B"] - 6) - 4, toe + 14
    z0, seen = 0.0, set()
    for k, (t_, nm) in enumerate(zip(p["Lt"], p["Ln"])):
        z1 = z0 + float(t_)
        ax.fill([xmin, xmax, xmax, xmin], [-z1, -z1, -z0, -z0], color=SOILS.get(nm, {}).get("col", "#e4d5b7"), alpha=0.75, label=None if nm in seen else f"Tanah {nm.lower()}")
        seen.add(nm)
        if k > 0: ax.plot([xmin, xmax], [-z0, -z0], color="#948065", lw=0.8, ls="--")
        ax.text(xmin + 0.5, -(z0 + z1) / 2, f"{nm} ({fmt(t_, 1)} m)", va="center", fontsize=8, color="#1e293b", fontweight='bold')
        z0 = z1
    ax.fill([xmin, xmin, p["B"], toe, xmin], [0, p["H"], p["H"], 0, 0], color="#f5eedc", ec="#785934", lw=1.6, label="Timbunan", alpha=0.9)
    ax.axhline(-p["D"], color="#64748b", lw=1.1, ls=":"); ax.text(xmax - 1, -p["D"] - 0.3, "tanah keras", ha="right", va="top", fontsize=8, color="#64748b")
    ax.plot([xmin, xmax], [0, 0], color="#785934", lw=1.2)
    if p["q"] > 0:
        ax.annotate("", xy=(xmin + 2, p["H"] + 0.05), xytext=(p["B"], p["H"] + 0.05), arrowprops=dict(arrowstyle="-", color="#dc2626"))
        for qx in np.linspace(xmin + 2, p["B"], 7):
            ax.annotate("", xy=(qx, p["H"]), xytext=(qx, p["H"] + 1.1), arrowprops=dict(arrowstyle="->", color="#dc2626", lw=0.8))
        ax.text((xmin + 2 + p["B"]) / 2, p["H"] + 1.4, f"q = {fmt(p['q'], 0)} kPa (lalin)", ha="center", fontsize=8, color="#dc2626", fontweight='bold')
    xs, ys = arc_points(xc, yc, R, r["xa"], r["xb"])
    ax.plot(xs, ys, color="#dc2626", lw=2.8, label=f"Bidang kritis FS={fmt(fs)}", zorder=5)
    for xv in np.linspace(r["xa"], r["xb"], 15):
        yb = yc - np.sqrt(max(R * R - (xv - xc) ** 2, 0))
        ax.plot([xv, xv], [yb, float(surface_y(xv, p))], color="#dc2626", lw=0.45, alpha=0.45)
    ax.plot([xc], [yc], marker="+", color="#dc2626", ms=12, mew=2.2, zorder=6)
    ax.plot([xc, r["xa"]], [yc, r["ya"]], color="#dc2626", lw=0.7, ls=":", alpha=0.7)
    ax.plot([xc, r["xb"]], [yc, r["yb_"]], color="#dc2626", lw=0.7, ls=":", alpha=0.7)
    ax.set_xlim(xmin, xmax); ax.set_ylim(-p["D"] - 1.5, max(yc + 1.0, p["H"] + 4)); ax.set_aspect("equal")
    ax.set_xlabel("Jarak dari sumbu jalan (m)", fontsize=9, fontweight='600'); ax.set_ylabel("Elevasi relatif (m)", fontsize=9, fontweight='600')
    ax.set_title("PENAMPANG & BIDANG LONGSOR KRITIS", fontsize=11, fontweight="bold")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.15), ncol=3, fontsize=8, frameon=False, labelcolor='#1e293b')
    ax.grid(alpha=0.35, color='#cbd5e1')
    X, Y = np.meshgrid(res["xcs"], res["ycs"])
    top = max(2.2 * fs, target * 1.3)
    G = np.ma.masked_invalid(np.minimum(res["grid"], top))
    levels = np.linspace(min(fs, target) * 0.98, top, 14)
    cs = ax2.contourf(X, Y, G, levels=levels, cmap="RdYlGn", vmin=levels[0], vmax=levels[-1])
    cb = fig.colorbar(cs, ax=ax2, label="FS minimum"); cb.ax.yaxis.label.set_color('#1e293b'); cb.ax.tick_params(colors='#334155')
    if np.nanmin(res["grid"]) < target < np.nanmax(G): ax2.contour(X, Y, G, levels=[target], colors="black", linewidths=1.3, linestyles="--")
    ax2.plot([xc], [yc], marker="*", color="black", ms=14, zorder=5)
    ax2.set_xlabel("Pusat x (m)", fontsize=9, fontweight='600'); ax2.set_ylabel("Pusat y (m)", fontsize=9, fontweight='600')
    ax2.set_title(f"PETA FS PUSAT (putus: FS={fmt(target)})", fontsize=10, fontweight="bold")
    ax2.grid(alpha=0.35, color='#cbd5e1'); fig.tight_layout()
    return fig, r

def dxf_bytes(p, res, r):
    fs, xc, yc, R = res["best"]
    doc = ezdxf.new("R2010", setup=True); doc.units = 6
    layers = [("TOL_TIMBUNAN", 8), ("TOL_TANAH_DASAR", 30), ("TOL_BIDANG_LONGSOR", 1), ("TOL_TEKS", 7),
              ("TOL_DIMENSI", 5), ("TOL_GRID", 253), ("TOL_IRISAN", 3), ("TOL_PUSAT", 6), ("TOL_BATAS_LAPISAN", 40)]
    for n, col in layers: doc.layers.add(n, color=col)
    msp = doc.modelspace()
    toe, xl, xr = p["B"] + p["m"] * p["H"], r["xa"] - 8, p["B"] + p["m"] * p["H"] + 18
    msp.add_lwpolyline([(xl, 0), (xl, p["H"]), (p["B"], p["H"]), (toe, 0)], close=True, dxfattribs={"layer": "TOL_TIMBUNAN", "lineweight": 35})
    try:
        hatch = msp.add_hatch(dxfattribs={"layer": "TOL_TIMBUNAN"}); hatch.set_pattern_fill("ANSI31", scale=0.5)
        hatch.paths.add_polyline_path([(xl, 0), (xl, p["H"]), (p["B"], p["H"]), (toe, 0)], is_closed=True)
    except Exception: pass
    z0 = 0.0
    for k, (t_, nm) in enumerate(zip(p["Lt"], p["Ln"])):
        z1 = z0 + float(t_)
        msp.add_lwpolyline([(xl, -z0), (xr, -z0), (xr, -z1), (xl, -z1)], close=True, dxfattribs={"layer": "TOL_TANAH_DASAR"})
        if k > 0: msp.add_line((xl, -z0), (xr, -z0), dxfattribs={"layer": "TOL_BATAS_LAPISAN", "linetype": "DASHED"})
        tx = msp.add_text(f"{nm} (t={t_:.1f}m, γ={p['Lg'][k]:.1f}, c={p['Lc'][k]:.0f}, φ={p['Lph'][k]:.0f}°)", height=0.4, dxfattribs={"layer": "TOL_TEKS"})
        tx.set_placement((xl + 0.5, -(z0 + z1) / 2)); z0 = z1
    msp.add_line((xl, -p["D"]), (xr, -p["D"]), dxfattribs={"layer": "TOL_TANAH_DASAR", "lineweight": 40})
    a0 = np.degrees(np.arctan2(r["ya"] - yc, r["xa"] - xc)) % 360
    a1 = np.degrees(np.arctan2(r["yb_"] - yc, r["xb"] - xc)) % 360
    msp.add_arc((float(xc), float(yc)), float(R), float(a0), float(a1), dxfattribs={"layer": "TOL_BIDANG_LONGSOR", "lineweight": 50})
    msp.add_circle((float(xc), float(yc)), 0.3, dxfattribs={"layer": "TOL_PUSAT"})
    msp.add_line((float(xc) - 0.5, float(yc)), (float(xc) + 0.5, float(yc)), dxfattribs={"layer": "TOL_PUSAT"})
    msp.add_line((float(xc), float(yc) - 0.5), (float(xc), float(yc) + 0.5), dxfattribs={"layer": "TOL_PUSAT"})
    msp.add_line((float(xc), float(yc)), (float(r["xa"]), float(r["ya"])), dxfattribs={"layer": "TOL_PUSAT", "linetype": "DASHDOT"})
    msp.add_line((float(xc), float(yc)), (float(r["xb"]), float(r["yb_"])), dxfattribs={"layer": "TOL_PUSAT", "linetype": "DASHDOT"})
    for xv in np.linspace(r["xa"], r["xb"], 20):
        yb_v = yc - np.sqrt(max(R * R - (xv - xc) ** 2, 0))
        msp.add_line((float(xv), float(yb_v)), (float(xv), float(surface_y(xv, p))), dxfattribs={"layer": "TOL_IRISAN", "linetype": "DASHED"})
    msp.add_line((float(p["B"]), float(p["H"]) + 1.5), (float(toe), float(p["H"]) + 1.5), dxfattribs={"layer": "TOL_DIMENSI"})
    tx_d = msp.add_text(f"m·H = {p['m']*p['H']:.1f} m", height=0.4, dxfattribs={"layer": "TOL_DIMENSI"}); tx_d.set_placement(((p["B"] + toe) / 2, p["H"] + 1.8))
    t1 = msp.add_text(f"TolSlope v2.0 | FS Bishop = {fs:.3f} | FS Fellenius = {r['fel']:.3f} | Pusat ({xc:.2f}; {yc:.2f}) R = {R:.2f} m", height=0.6, dxfattribs={"layer": "TOL_TEKS"})
    t1.set_placement((xl, p["H"] + 3.2))
    t2 = msp.add_text(f"Timbunan: H={p['H']}m, B={p['B']}m, m={p['m']}, γ={p['gE']}kN/m³, c'={p['cE']}kPa, φ'={p['phE']}°, q={p['q']}kPa, ru={p['ru']} | {datetime.now():%d-%m-%Y %H:%M}", height=0.35, dxfattribs={"layer": "TOL_TEKS"})
    t2.set_placement((xl, p["H"] + 2.5))
    buf = io.StringIO(); doc.write(buf)
    return buf.getvalue().encode("utf-8")

def excel_bytes(P0, P, res, r, det, sens, target, png):
    fs, xc, yc, R = res["best"]
    A = "Arial"
    fn, fb = Font(name=A, size=10), Font(name=A, size=10, bold=True)
    fin, fh = Font(name=A, size=10, color="0000FF"), Font(name=A, size=10, bold=True, color="FFFFFF")
    fill_h, fill_s = PatternFill("solid", fgColor="1F3A5F"), PatternFill("solid", fgColor="DCE6F1")
    fill_in, fill_res = PatternFill("solid", fgColor="FFFFDD"), PatternFill("solid", fgColor="E8F5E9")
    thin = Side(style="thin", color="BFBFBF"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
    cen, left_w = Alignment(horizontal="center", vertical="center", wrap_text=True), Alignment(horizontal="left", vertical="center", wrap_text=True)
    wb = Workbook()
    stamp = f"TolSlope v2.0 · Fellenius & Bishop Disederhanakan · Dibuat {datetime.now():%d-%m-%Y %H:%M}"

    def put(ws, row, col, val, font=fn, fmt_=None, al=None, fill=None):
        c = ws.cell(row=row, column=col, value=val); c.font = font
        if fmt_: c.number_format = fmt_
        if al: c.alignment = al
        c.border = box
        if fill: c.fill = fill
        return c

    def title(ws, text, sub):
        ws["A1"], ws["A2"] = text, sub
        ws["A1"].font, ws["A2"].font = Font(name=A, size=15, bold=True, color="1F3A5F"), Font(name=A, size=9, italic=True, color="595959")
        ws.sheet_view.showGridLines = False; ws.page_setup.orientation = "landscape"
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    def section(ws, row, text, n):
        for j in range(1, n + 1): ws.cell(row=row, column=j).fill = fill_s
        ws.cell(row=row, column=1, value=text).font = fb

    def header(ws, row, labels):
        for j, t in enumerate(labels, 1):
            c = put(ws, row, j, t, font=fh, al=cen); c.fill = fill_h
        ws.row_dimensions[row].height = 28

    ok_s = dict(fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), font=Font(name=A, bold=True, color="006100"))
    bad_s = dict(fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"), font=Font(name=A, bold=True, color="9C0006"))

    # Sheet 1: Ringkasan
    ws = wb.active; ws.title = "Ringkasan"
    title(ws, "LAPORAN ANALISIS STABILITAS LERENG TIMBUNAN JALAN TOL", stamp)
    for col, w in zip("ABCDEFGH", [38, 20, 14, 38, 12, 12, 12, 12]): ws.column_dimensions[col].width = w
    row = 4; section(ws, row, "▶ HASIL ANALISIS", 4)
    header(ws, row + 1, ["Parameter", "Nilai", "Satuan", "Keterangan & Asal Data"])
    r_fs, r_fel, r_t, r_st = row + 2, row + 3, row + 4, row + 5
    put(ws, r_fs, 1, "FS Bishop (lingkaran kritis)", fb); put(ws, r_fs, 2, float(fs), fb, "0.000", cen, fill_res); put(ws, r_fs, 3, "–", al=cen); put(ws, r_fs, 4, "Hasil iterasi Bishop disederhanakan (kriteria utama)", al=left_w)
    put(ws, r_fel, 1, "FS Fellenius (lingkaran sama)"); put(ws, r_fel, 2, float(r["fel"]), fn, "0.000", cen, fill_res); put(ws, r_fel, 3, "–", al=cen); put(ws, r_fel, 4, "Metode Fellenius (konservatif tanpa iterasi)", al=left_w)
    put(ws, r_t, 1, "FS disyaratkan (SNI 8460:2017)"); put(ws, r_t, 2, float(target), fin, "0.00", cen, fill_in); put(ws, r_t, 3, "–", al=cen); put(ws, r_t, 4, "Input pengguna (ref: SNI 8460:2017 statis ≥ 1.50)", al=left_w)
    put(ws, r_st, 1, "Status Kestabilan", fb); put(ws, r_st, 2, f'=IF(B{r_fs}>=B{r_t},"MEMENUHI","TIDAK MEMENUHI")', fb, None, cen); put(ws, r_st, 3, "–", al=cen); put(ws, r_st, 4, f'=IF(B{r_fs}>=B{r_t},"Lereng stabil aman","Lereng rawan longsor")', al=left_w)
    ws.conditional_formatting.add(f"B{r_st}", FormulaRule(formula=[f'$B${r_st}="MEMENUHI"'], **ok_s))
    ws.conditional_formatting.add(f"B{r_st}", FormulaRule(formula=[f'$B${r_st}="TIDAK MEMENUHI"'], **bad_s))

    row = r_st + 2; section(ws, row, "▶ DATA MASUKAN TIMBUNAN, BEBAN, DAN AIR PORI", 4)
    header(ws, row + 1, ["Parameter", "Nilai", "Satuan", "Asal Data / Standar Pengujian"])
    items = [("Tinggi timbunan H", P0["H"], "m", "0.0", "Gambar Desain Geometri Jalan Tol (DED)"),
             ("Lebar puncak B", P0["B"], "m", "0.0", "Desain Penampang Melintang Jalan Tol"),
             ("Kemiringan 1 : m (V:H)", P0["m"], "–", "0.00", "Spesifikasi Lereng Badan Jalan Bina Marga"),
             ("Berat isi timbunan γ", P0["gE"], "kN/m³", "0.0", "Uji Pemadatan Lab (Proctor SNI 1742)"),
             ("Kohesi timbunan c'", P0["cE"], "kPa", "0.0", "Uji Triaksial Efektif CD/CU (SNI 2813)"),
             ("Sudut geser timbunan φ'", P0["phE"], "°", "0.0", "Uji Triaksial Efektif / Direct Shear"),
             ("Beban lalu lintas q", P0["q"], "kPa", "0.0", "Beban Kendaraan Berat Ekuivalen (Bina Marga)"),
             ("Rasio air pori ru", P0["ru"], "–", "0.00", "Piezometer Lapangan / Model Rembesan")]
    for k, (nm, val, un, f_, asal) in enumerate(items):
        put(ws, row + 2 + k, 1, nm); put(ws, row + 2 + k, 2, float(val), fin, f_, cen, fill_in); put(ws, row + 2 + k, 3, un, al=cen); put(ws, row + 2 + k, 4, asal, al=left_w)

    row = row + 2 + len(items) + 1; section(ws, row, "▶ DATA TANAH DASAR", 8)
    header(ws, row + 1, ["Lapisan", "Tebal (m)", "Kedalaman Atas", "Kedalaman Bawah", "γ (kN/m³)", "c (kPa)", "φ (°)", "Asal Data Penyelidikan"])
    for col, w in zip("EFGH", [12, 12, 12, 30]): ws.column_dimensions[col].width = w
    first = row + 2
    for k, (nm, t_, g_, c_, p_) in enumerate(zip(P["Ln"], P["Lt"], P["Lg"], P["Lc"], P["Lph"])):
        rr = first + k
        put(ws, rr, 1, f"Lapisan {k+1}: {nm}"); put(ws, rr, 2, float(t_), fin, "0.0", cen, fill_in)
        put(ws, rr, 3, 0 if k == 0 else f"=D{rr-1}", fn, "0.0", cen); put(ws, rr, 4, f"=C{rr}+B{rr}", fn, "0.0", cen)
        put(ws, rr, 5, float(g_), fin, "0.0", cen, fill_in); put(ws, rr, 6, float(c_), fin, "0.0", cen, fill_in); put(ws, rr, 7, float(p_), fin, "0.0", cen, fill_in)
        put(ws, rr, 8, "Uji Bor Mesin (SPT ASTM D1586) & Lab UDS", al=left_w)

    # Sheet 2: Irisan
    wi = wb.create_sheet("Irisan"); title(wi, "TABEL IRISAN BIDANG GELINCIR KRITIS", f"{len(det)} irisan · FS Bishop = {fs:.3f} · Pusat ({xc:.2f}; {yc:.2f}) · R = {R:.2f} m")
    for col, w in zip("ABCDEFGHI", [8, 16, 14, 16, 12, 12, 12, 16, 16]): wi.column_dimensions[col].width = w
    header(wi, 4, ["No", "x tengah (m)", "Lebar b (m)", "Berat W (kN/m)", "α (°)", "c (kPa)", "φ (°)", "W sin α (kN/m)", "c·b (kN/m)"])
    cols_d = ["x tengah (m)", "lebar b (m)", "berat W (kN/m)", "α (°)", "c (kPa)", "φ (°)", "W sin α", "c·b (kN/m)"]
    for k, rowv in enumerate(det[cols_d].itertuples(index=False)):
        rr = 5 + k; put(wi, rr, 1, k + 1, al=cen)
        for j in range(len(cols_d)): put(wi, rr, 2 + j, float(rowv[j]), fn, "0.00" if j < 2 else "0.0", cen)
    e = 4 + len(det); put(wi, e + 1, 1, "Σ", fb, al=cen); put(wi, e + 1, 4, f"=SUM(D5:D{e})", fb, "0.0", cen); put(wi, e + 1, 8, f"=SUM(H5:H{e})", fb, "0.0", cen)
    wi.freeze_panes = "A5"

    # Sheet 3: Sensitivitas
    wsn = wb.create_sheet("Sensitivitas"); title(wsn, "SENSITIVITAS KEMIRINGAN LERENG TERHADAP FS", "Evaluasi kemiringan lereng 1:m terhadap pemenuhan kriteria SNI")
    for col, w in zip("ABCDE", [16, 14, 14, 14, 18]): wsn.column_dimensions[col].width = w
    header(wsn, 4, ["Lereng 1:m", "FS Bishop", "Target FS", "Selisih", "Status"])
    n = len(sens)
    for k, (mm, fv) in enumerate(zip(sens["m (1 : m)"], sens["FS Bishop"])):
        rr = 5 + k
        put(wsn, rr, 1, float(mm), fn, "0.0", cen); put(wsn, rr, 2, None if pd.isna(fv) else float(fv), fn, "0.000", cen)
        put(wsn, rr, 3, f"=Ringkasan!$B${r_t}", fn, "0.00", cen); put(wsn, rr, 4, f'=IF(B{rr}="","",B{rr}-C{rr})', fn, "+0.000;-0.000;0.000", cen)
        put(wsn, rr, 5, f'=IF(B{rr}="","-",IF(B{rr}>=C{rr},"Memenuhi","Tidak memenuhi"))', fn, None, cen)
    ch = LineChart(); ch.title, ch.height, ch.width = "FS Bishop vs Kemiringan Lereng", 9, 16
    ch.y_axis.title, ch.x_axis.title = "FS", "Kemiringan 1 : m"
    ch.add_data(Reference(wsn, min_col=2, max_col=3, min_row=4, max_row=4 + n), titles_from_data=True)
    ch.set_categories(Reference(wsn, min_col=1, min_row=5, max_row=4 + n)); wsn.add_chart(ch, "G4")

    # Sheet 4: Penampang
    wp = wb.create_sheet("Penampang"); title(wp, "GAMBAR PENAMPANG BIDANG GELINCIR KRITIS", f"FS Bishop = {fs:.3f}")
    img = XLImage(io.BytesIO(png)); k_ = 1000 / img.width; img.width, img.height = int(img.width * k_), int(img.height * k_)
    wp.add_image(img, "A4")

    # Sheet 5: Referensi Rumus
    wr = wb.create_sheet("Referensi Rumus"); title(wr, "DOKUMENTASI RUMUS & STANDAR GEOTEKNIK", "Metode Keseimbangan Batas Fellenius, Bishop, dan SNI 8460:2017")
    wr.column_dimensions["A"].width, wr.column_dimensions["B"].width = 45, 55
    ref_rows = [
        ("1. METODE FELLENIUS (1927)", "FS = Σ[c·b/cos α + (W·cos α - u·b/cos α)·tan φ] / Σ(W·sin α)"),
        ("2. METODE BISHOP DISEDERHANAKAN (1955)", "FS = Σ[(c·b + (W - u·b)·tan φ) / mα] / Σ(W·sin α)"),
        ("Faktor Koreksi mα", "mα = cos α + (sin α · tan φ) / FS (dihitung iterasi konvergen)"),
        ("3. TEKANAN AIR PORI (ru)", "ru = u / (γ·z) -> u·b = ru · W_tanah"),
        ("4. STANDAR FAKTOR KEAMANAN (SNI 8460:2017)", "Statis Jangka Panjang FS ≥ 1.50 | Jangka Pendek FS ≥ 1.30 | Gempa FS ≥ 1.10"),
    ]
    for idx_r, (c1, c2) in enumerate(ref_rows):
        r_ = 4 + idx_r
        put(wr, r_, 1, c1, fb, al=left_w); put(wr, r_, 2, c2, fn, al=left_w)
    out = io.BytesIO(); wb.save(out); return out.getvalue()

# ============================================================================
# ANIMASI HTML5 / JS CANVAS (TAMPILAN CERAH / DAYLIGHT THEME)
# ============================================================================
ANIM_TEMPLATE = r"""<!DOCTYPE html><html><head><meta charset="utf-8"><style>
body{margin:0;font:12px 'Plus Jakarta Sans',Arial,sans-serif;color:#1e293b;background:#f8fafc}
.bar{display:flex;gap:6px;align-items:center;padding:6px 2px}
button,select{font:inherit;padding:5px 12px;border:1px solid #cbd5e1;border-radius:6px;background:#ffffff;color:#1e293b;cursor:pointer;font-weight:600}
button.on{background:linear-gradient(135deg,#1d4ed8,#0284c7);color:#fff;border-color:transparent}
canvas{width:100%;display:block;border:1px solid #cbd5e1;border-radius:10px;background:#eef6ff}
input[type=range]{flex:1;min-width:100px;accent-color:#0284c7}
</style></head><body>
<div class="bar"><button id="m1" class="on">1 · Pencarian Kritis</button><button id="m2">2 · Simulasi Longsor</button><span style="flex:1"></span>
<button id="pp">⏸ Jeda</button><button id="rs">⟲ Ulang</button><select id="sp"><option value="0.5">0.5×</option><option value="1" selected>1×</option><option value="2">2×</option></select></div>
<canvas id="cv"></canvas>
<div class="bar"><input type="range" id="sl" min="0" max="1000" value="0"></div>
<script>
const D=__DATA__;const cv=document.getElementById('cv'),ctx=cv.getContext('2d');
let mode=1,playing=true,speed=1,t=0,last=null,W=0,Hh=0,S=1,ox=0,oy=0,drag=false;
const T1=14,T2=9,TR=0.35,N=D.trials.length;
const clamp=(v,a,b)=>Math.max(a,Math.min(b,v)),lerp=(a,b,u)=>a+(b-a)*u,ease=u=>u*u*(3-2*u);
const order=D.trials.map((c,i)=>i).sort((a,b)=>D.trials[a].fs-D.trials[b].fs);
const rank=new Array(N);order.forEach((idx,k)=>rank[idx]=k/Math.max(1,N-1));
const thMax=clamp(0.18+0.8*(1-D.fs),0.2,0.5);
function rot(p,th){const dx=p[0]-D.xc,dy=p[1]-D.yc,c=Math.cos(th),s=Math.sin(th);return [D.xc+dx*c-dy*s,D.yc+dx*s+dy*c]}
let BD=[D.xmin,D.xmax,-D.D-1.2,Math.max(D.yc+2,D.H+3)];
if(D.fs<1){D.mass.forEach(p=>{const q=rot(p,thMax);BD[1]=Math.max(BD[1],q[0]+2);BD[0]=Math.min(BD[0],q[0]-2);BD[3]=Math.max(BD[3],q[1]+1)})}
const sx=x=>ox+S*x,sy=y=>oy-S*y;
function layout(){const dpr=window.devicePixelRatio||1,w=cv.parentElement.clientWidth,wr=BD[1]-BD[0],hr=BD[3]-BD[2];
 const h=Math.round(Math.min(w*hr/wr,480));cv.style.height=h+'px';cv.width=w*dpr;cv.height=h*dpr;ctx.setTransform(dpr,0,0,dpr,0,0);
 W=w;Hh=h;S=Math.min(w/wr,h/hr);ox=(w-S*wr)/2-S*BD[0];oy=(h+S*hr)/2+S*BD[2]}
function addPoly(pts){pts.forEach((p,i)=>i?ctx.lineTo(sx(p[0]),sy(p[1])):ctx.moveTo(sx(p[0]),sy(p[1])));ctx.closePath()}
function strokeArc(pts,prog,col,lw){const n=Math.max(2,Math.round(pts.length*prog));ctx.beginPath();pts.slice(0,n).forEach((p,i)=>i?ctx.lineTo(sx(p[0]),sy(p[1])):ctx.moveTo(sx(p[0]),sy(p[1])));ctx.strokeStyle=col;ctx.lineWidth=lw;ctx.stroke()}
function terrain(){const x0=sx(BD[0]-6),ww=sx(BD[1]+6)-x0;
 ctx.fillStyle='#d5dbe3';ctx.fillRect(x0,sy(-D.D),ww,sy(BD[2]-2)-sy(-D.D));
 D.layers.forEach(L=>{ctx.fillStyle=L.col;ctx.globalAlpha=0.8;ctx.fillRect(x0,sy(-L.z0),ww,sy(-L.z1)-sy(-L.z0));ctx.globalAlpha=1});
 ctx.beginPath();addPoly(D.emb);ctx.fillStyle='#f5eedc';ctx.globalAlpha=0.92;ctx.fill();ctx.globalAlpha=1;ctx.strokeStyle='#785934';ctx.lineWidth=1.3;ctx.stroke();
 ctx.beginPath();ctx.moveTo(x0,sy(0));ctx.lineTo(x0+ww,sy(0));ctx.strokeStyle='#785934';ctx.lineWidth=1;ctx.stroke()}
function fsCol(fs,a){return 'hsla('+120*clamp((fs-1)/(Math.max(D.target*1.3,1.3)-1),0,1)+',80%,42%,'+a+')'}
function drawM1(){
 ctx.fillStyle='#eef6ff';ctx.fillRect(0,0,W,Hh);terrain();
 const ph2=clamp((t-7)/3,0,1),ph3=clamp((t-10)/3,0,1),keep=lerp(1,0.15,ease(ph2))-ph3*0.5;
 for(let i=0;i<N;i++){const ta=0.5+i/N*6.2;if(t<ta)continue;const a=0.6*clamp(1-(rank[i]-keep)*12,0,1);if(a<0.02)continue;strokeArc(D.trials[i].pts,clamp((t-ta)/TR,0,1),fsCol(D.trials[i].fs,a),1.5)}
 const u=clamp((t-9.5)/1.5,0,1);
 if(u>0){strokeArc(D.crit.pts,1,'rgba(220,38,38,'+u+')',lerp(2,4.5,u));ctx.fillStyle='#dc2626';ctx.beginPath();ctx.arc(sx(D.xc),sy(D.yc),4.5,0,7);ctx.fill()}}
function drawM2(){
 let th=D.fs<1?thMax*ease(clamp((t-1)/5,0,1)):(D.fs<D.target?0.015*Math.sin(t*8):0);
 ctx.fillStyle='#eef6ff';ctx.fillRect(0,0,W,Hh);
 ctx.save();ctx.beginPath();ctx.rect(0,0,W,Hh);addPoly(D.mass);ctx.clip('evenodd');terrain();ctx.restore();
 ctx.save();ctx.translate(sx(D.xc),sy(D.yc));ctx.rotate(-th);ctx.translate(-sx(D.xc),-sy(D.yc));
 ctx.beginPath();addPoly(D.mass);ctx.save();ctx.clip();terrain();ctx.restore();ctx.strokeStyle='#785934';ctx.lineWidth=1.3;ctx.stroke();ctx.restore();
 strokeArc(D.crit.pts,1,'#dc2626',2.6)}
function render(){mode===1?drawM1():drawM2()}
function loop(ts){if(last===null)last=ts;const dt=Math.min(0.05,(ts-last)/1000);last=ts;
 const T=mode===1?T1:T2;if(playing&&!drag){t+=dt*speed;if(t>=T)t-=T}
 render();if(!drag)document.getElementById('sl').value=Math.round(t/T*1000);requestAnimationFrame(loop)}
function setMode(m){mode=m;t=0;document.getElementById('m1').className=m===1?'on':'';document.getElementById('m2').className=m===2?'on':''}
document.getElementById('m1').onclick=()=>setMode(1);document.getElementById('m2').onclick=()=>setMode(2);
document.getElementById('pp').onclick=e=>{playing=!playing;e.target.textContent=playing?'⏸ Jeda':'▶ Putar'};
document.getElementById('rs').onclick=()=>{t=0};document.getElementById('sp').onchange=e=>{speed=parseFloat(e.target.value)};
const sl=document.getElementById('sl');sl.oninput=()=>{drag=true;t=sl.value/1000*(mode===1?T1:T2)};sl.onchange=()=>{drag=false};
window.addEventListener('resize',layout);layout();setMode(1);requestAnimationFrame(loop);
</script></body></html>"""

def trial_circles(p, best, n_far=35, n_near=12, seed=7):
    fs0, xc0, yc0, R0 = best; H, m, B, D = p["H"], p["m"], p["B"], p["D"]; toe = B + m * H; rng = np.random.default_rng(seed)
    out, far, near, tries = [], 0, 0, 0
    def add(xc, yc, ybot):
        R = yc - ybot
        if R < 1.0: return False
        r = circle_fs(p, xc, yc, R)
        if r is None: return False
        xs, ys = arc_points(xc, yc, R, r["xa"], r["xb"], 32)
        out.append(dict(fs=round(float(r["fs"]), 3), pts=[[round(float(a), 2), round(float(b), 2)] for a, b in zip(xs, ys)]))
        return True
    while (far < n_far or near < n_near) and tries < 2000:
        tries += 1
        if far < n_far: far += add(rng.uniform(B - 1.5 * H, toe + 0.5 * H), rng.uniform(0.9 * H, 2.5 * H + 0.8 * D), rng.uniform(-D, 0.8 * H))
        if near < n_near: near += add(xc0 + rng.normal(0, 0.6 * max(1.0, 0.6 * H)), yc0 + rng.normal(0, 0.6 * max(1.0, 0.6 * H)), float(np.clip((yc0 - R0) + rng.normal(0, 0.8), -D, 0.8 * H)))
    return [out[i] for i in rng.permutation(len(out))]

@st.cache_data(show_spinner="🎬 Menyiapkan animasi...")
def cached_trials(pt, best): return trial_circles(prep(dict(zip(PKEYS, pt))), best)

def anim_html(p, res, r, target, trials):
    fs, xc, yc, R = res["best"]; toe = p["B"] + p["m"] * p["H"]
    rd = lambda a: [[round(float(x), 2), round(float(y), 2)] for x, y in a]
    xa, ya = arc_points(xc, yc, R, r["xa"], r["xb"], 50); sxs = np.linspace(r["xb"], r["xa"], 50)
    layers, z0 = [], 0.0
    for t_, nm in zip(p["Lt"], p["Ln"]):
        layers.append(dict(z0=z0, z1=z0 + float(t_), t=float(t_), name=nm, col=SOILS.get(nm, {}).get("col", "#e4d5b7"))); z0 += float(t_)
    xmin = min(r["xa"], p["B"] - 6) - 4
    data = dict(H=float(p["H"]), D=float(p["D"]), xmin=float(xmin), xmax=float(toe + 14), xc=float(xc), yc=float(yc), fs=float(fs), target=float(target),
                layers=layers, emb=rd([(xmin, 0), (xmin, p["H"]), (p["B"], p["H"]), (toe, 0)]),
                crit=dict(pts=rd(zip(xa, ya))), mass=rd(list(zip(xa, ya)) + list(zip(sxs, surface_y(sxs, p)))), trials=trials)
    return ANIM_TEMPLATE.replace("__DATA__", json.dumps(data, ensure_ascii=False))

# ============================================================================
# ANTARMUKA STREAMLIT
# ============================================================================
st.markdown('<div class="hero-title">🏗️ TolSlope v2.0</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-sub">Analisis Stabilitas Lereng Timbunan Jalan Tol — Fellenius & Bishop Disederhanakan (SNI 8460:2017)</div>', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## 🎛️ Parameter Lereng")
    nl = st.radio("Jumlah lapisan tanah dasar", [1, 2, 3], horizontal=True)
    types = [st.selectbox(f"Jenis Lapisan {i + 1}" + (" (atas)" if i == 0 else ""), list(SOILS), index=i, key=f"j_{i}") for i in range(nl)]
    with st.form("param"):
        st.markdown("### 🏗️ Timbunan")
        H = st.number_input("Tinggi H (m)", 1.0, 15.0, 5.0, 0.5)
        B = st.number_input("Lebar puncak B (m)", 3.0, 40.0, 16.8, 0.1)
        m = st.number_input("Lereng 1 : m (V:H)", 1.0, 4.0, 2.0, 0.25)
        gE = st.number_input("Berat isi γ (kN/m³)", 14.0, 22.0, 18.0, 0.5)
        cE = st.number_input("Kohesi c' (kPa)", 0.0, 50.0, 5.0, 1.0)
        phE = st.number_input("Sudut geser φ' (°)", 0.0, 45.0, 30.0, 1.0)
        P0 = dict(H=H, B=B, m=m, gE=gE, cE=cE, phE=phE)
        for i in (1, 2, 3):
            if i <= nl:
                tp = types[i - 1]; sd = SOILS[tp]
                st.markdown(f"### 🌍 Lapisan {i}: {tp}")
                P0[f"t{i}"] = st.number_input(f"Tebal (m) L{i}", 0.5, 30.0, sd["t"], 0.5)
                P0[f"g{i}"] = st.number_input(f"γ (kN/m³) L{i}", 12.0, 22.0, sd["g"], 0.5)
                P0[f"c{i}"] = st.number_input(f"c / cu (kPa) L{i}", 0.0, 300.0, sd["c"], 1.0)
                P0[f"ph{i}"] = st.number_input(f"φ (°) L{i}", 0.0, 45.0, sd["ph"], 1.0)
            else: P0.update({f"t{i}": 0.0, f"g{i}": 16.0, f"c{i}": 0.0, f"ph{i}": 0.0})
        st.markdown("### 🚗 Beban & Air Pori")
        P0["q"] = st.number_input("Beban lalu lintas q (kPa)", 0.0, 40.0, 10.0, 1.0)
        P0["ru"] = st.slider("Rasio air pori ru", 0.0, 0.6, 0.0, 0.05)
        st.markdown("### 🎯 Kriteria Keamanan")
        target = st.number_input("FS Minimum Disyaratkan", 1.0, 2.5, 1.5, 0.05)
        fine = st.checkbox("🔬 Grid pencarian rapat", value=False)
        st.form_submit_button("🔍 HITUNG STABILITAS", use_container_width=True)

PT = tuple(float(P0[k]) for k in PKEYS)
P = prep(P0, types)
grid_cfg = (22, 16, 28, True) if fine else (15, 11, 20, True)

def sensitivity(P0):
    rows = []
    for mm in [1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0]:
        pt = tuple(float(mm) if k == "m" else float(P0[k]) for k in PKEYS)
        rr = cached_search(pt, 11, 8, 14, False)
        rows.append((mm, rr["best"][0] if rr["best"] else np.nan))
    return pd.DataFrame(rows, columns=["m (1 : m)", "FS Bishop"])

sens = sensitivity(P0)
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📉 Stabilitas Lereng", "📈 Sensitivitas Kemiringan", "🎬 Animasi", "📐 Rumus & Teori", "📋 Asal Data"])

# TAB 1: STABILITAS
with tab1:
    res = cached_search(PT, *grid_cfg)
    if res["best"] is None:
        st.error("❌ Tidak ada bidang gelincir valid. Sesuaikan parameter lereng.")
    else:
        fs, xc, yc, R = res["best"]; fig, r = plot_result(P, res, target); ok = fs >= target
        badge_cls = "badge-safe" if ok else ("badge-warn" if fs >= 1.0 else "badge-danger")
        badge_txt = "✅ LERENG AMAN MEMENUHI SNI 8460:2017" if ok else ("⚠️ LERENG RAWAN (FS < TARGET)" if fs >= 1.0 else "❌ LERENG TIDAK STABIL (POTENSI LONGSOR)")
        st.markdown(f'<div style="text-align:center;margin:0.8rem 0;"><span class="status-badge {badge_cls}">{badge_txt}</span></div>', unsafe_allow_html=True)
        col_c = "safe" if ok else ("warn" if fs >= 1.0 else "danger")
        st.markdown(f"""<div class="metric-row">
            <div class="metric-card"><div class="metric-label">FS Bishop (Kritis)</div><div class="metric-val {col_c}">{fmt(fs,3)}</div><span class="data-origin origin-calc">Iterasi Bishop</span></div>
            <div class="metric-card"><div class="metric-label">FS Fellenius</div><div class="metric-val neutral">{fmt(r['fel'],3)}</div><span class="data-origin origin-calc">Metode 1927</span></div>
            <div class="metric-card"><div class="metric-label">Target Syarat</div><div class="metric-val neutral">{fmt(target)}</div><span class="data-origin origin-ref">SNI 8460:2017</span></div>
            <div class="metric-card"><div class="metric-label">Margin ΔFS</div><div class="metric-val {col_c}">{'+' if fs>=target else ''}{fmt(fs-target,3)}</div><span class="data-origin origin-calc">FS - Target</span></div>
        </div>""", unsafe_allow_html=True)
        st.markdown(f"""<div class="info-card">
            <h3>⚖️ Keseimbangan Gaya & Momen (Limit Equilibrium Breakdown)</h3>
            <table style="width:100%;font-size:0.9rem;border-collapse:collapse">
                <tr style="border-bottom:1px solid #e2e8f0"><td style="padding:7px;color:#64748b;font-weight:600">Total Gaya Dorong Gravitasi (Σ W sin α)</td><td style="padding:7px;font-weight:700;color:#dc2626">{fmt(r['drive'],1)} kN/m</td><td style="color:#475569">Massa lereng pemicu luncuran</td></tr>
                <tr style="border-bottom:1px solid #e2e8f0"><td style="padding:7px;color:#64748b;font-weight:600">Total Kuat Geser Penahan Bishop (Σ Kuat Geser)</td><td style="padding:7px;font-weight:700;color:#059669">{fmt(r['resist_bishop'],1)} kN/m</td><td style="color:#475569">Mobilisasi kohesi c dan friksi tan φ</td></tr>
                <tr style="border-bottom:1px solid #e2e8f0"><td style="padding:7px;color:#64748b;font-weight:600">Total Berat Massa Tanah Terlibat (Σ W)</td><td style="padding:7px;font-weight:700;color:#0284c7">{fmt(r['total_w'],1)} kN/m</td><td style="color:#475569">Akumulasi timbunan + beban q + tanah dasar</td></tr>
                <tr><td style="padding:7px;color:#64748b;font-weight:600">Rasio Faktor Keamanan (FS = Penahan / Pendorong)</td><td style="padding:7px;font-weight:800;color:{'#059669' if ok else '#dc2626'}">{fmt(r['resist_bishop'],1)} / {fmt(r['drive'],1)} = {fmt(fs,3)}</td><td style="color:#475569">Target SNI 8460: FS ≥ {fmt(target)}</td></tr>
            </table>
        </div>""", unsafe_allow_html=True)
        png = io.BytesIO(); fig.savefig(png, format="png", dpi=120, bbox_inches="tight", facecolor='#ffffff'); st.pyplot(fig); plt.close(fig)
        st.markdown(f"""<div class="info-card"><h3>🎯 Parameter Busur Kritis</h3>
        Pusat lingkaran: <strong>(x = {fmt(xc)} m; y = {fmt(yc)} m)</strong> | Jari-jari: <strong>R = {fmt(R)} m</strong> | Titik masuk: <strong>x = {fmt(r['xa'])} m</strong> | Titik keluar: <strong>x = {fmt(r['xb'])} m</strong>
        </div>""", unsafe_allow_html=True)
        det = (circle_fs(P, xc, yc, R, nsl=30, detail=True) or circle_fs(P, xc, yc, R, detail=True))["slices"]
        c1, c2, c3 = st.columns(3)
        c1.download_button("📊 Unduh Laporan Excel (.xlsx)", excel_bytes(P0, P, res, r, det, sens, target, png.getvalue()), file_name="laporan_stabilitas_lereng.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        c2.download_button("📐 Unduh Gambar CAD (.dxf)", dxf_bytes(P, res, r), file_name="stabilitas_lereng_tol.dxf", mime="application/dxf", use_container_width=True)
        c3.download_button("📄 Unduh Irisan (.csv)", det.to_csv(index=False).encode("utf-8"), file_name="irisan_lingkaran_kritis.csv", mime="text/csv", use_container_width=True)
        with st.expander(f"📊 Tabel {len(det)} Irisan Busur Kritis"): st.dataframe(det, hide_index=True, use_container_width=True)

# TAB 2: SENSITIVITAS
with tab2:
    st.markdown('<div class="info-card"><h3>📈 Pengaruh Kemiringan Lereng terhadap Faktor Keamanan</h3>Variasi kemiringan lereng 1:m (V:H) untuk menentukan lereng paling curam yang memenuhi syarat.</div>', unsafe_allow_html=True)
    fig2, ax3 = plt.subplots(figsize=(9, 4)); fig2.patch.set_facecolor('#ffffff'); ax3.set_facecolor('#f8fafc')
    cols_b = ['#059669' if v >= target else '#dc2626' for v in sens["FS Bishop"]]
    ax3.bar(sens["m (1 : m)"].astype(str), sens["FS Bishop"], color=cols_b, alpha=0.75, width=0.45)
    ax3.plot(sens["m (1 : m)"].astype(str), sens["FS Bishop"], marker="o", color="#0284c7", lw=2.2)
    ax3.axhline(target, color="#dc2626", ls="--", lw=1.8, label=f"Target FS = {fmt(target)}")
    ax3.set_xlabel("Kemiringan Lereng 1 : m", color='#1e293b', fontweight='600'); ax3.set_ylabel("FS Bishop", color='#1e293b', fontweight='600'); ax3.tick_params(colors='#334155')
    for s in ax3.spines.values(): s.set_color('#cbd5e1')
    ax3.legend(labelcolor='#1e293b', framealpha=0.7); ax3.grid(alpha=0.35, color='#cbd5e1'); fig2.tight_layout(); st.pyplot(fig2); plt.close(fig2)
    okm = sens[sens["FS Bishop"] >= target]
    if len(okm): st.success(f"✅ Kemiringan paling curam yang memenuhi target FS ≥ {fmt(target)} adalah **1 : {fmt(okm['m (1 : m)'].min(), 1)}**.")
    else: st.warning("⚠️ Tidak ada kemiringan pada variasi ini yang memenuhi target. Perlu perbaikan tanah dasar atau geotekstil.")
    st.dataframe(sens.round(3), hide_index=True, use_container_width=True)

# TAB 3: ANIMASI
with tab3:
    if res["best"] is None: st.info("Belum ada bidang longsor valid untuk dianimasikan.")
    else:
        st.markdown('<div class="info-card"><h3>🎬 Animasi Interaktif: Pencarian & Gelincir Lereng</h3>Mode 1 menunjukkan uji lingkaran komputasi; Mode 2 menunjukkan simulasi rotasi massa tanah.</div>', unsafe_allow_html=True)
        trials = cached_trials(PT, res["best"]); components.html(anim_html(P, res, r, target, trials), height=640, scrolling=False)

# TAB 4: RUMUS & TEORI
with tab4:
    st.markdown('<div class="info-card"><h3>📐 Formulasi Matematis Metode Keseimbangan Batas (LEM)</h3></div>', unsafe_allow_html=True)
    c_f, c_b = st.columns(2)
    with c_f:
        st.markdown("#### 🔵 Metode Fellenius (Ordinary Method, 1927)")
        st.latex(r"FS_{\text{Fellenius}} = \frac{\sum \left[ \frac{c' \cdot b}{\cos \alpha} + \left(W \cos \alpha - \frac{u \cdot b}{\cos \alpha}\right) \tan \phi' \right]}{\sum W \sin \alpha}")
        st.caption("Memenuhi keseimbangan momen rotasi. Mengabaikan gaya antar-irisan. Bersifat konservatif.")
    with c_b:
        st.markdown("#### 🟢 Metode Bishop Disederhanakan (1955)")
        st.latex(r"FS_{\text{Bishop}} = \frac{1}{\sum W \sin \alpha} \sum \left[ \frac{c' \cdot b + (W - u \cdot b) \tan \phi'}{m_\alpha} \right]")
        st.latex(r"m_\alpha = \cos \alpha + \frac{\sin \alpha \cdot \tan \phi'}{FS}")
        st.caption("Memenuhi keseimbangan momen dan gaya vertikal. Dihitung melalui iterasi konvergensi numerik.")
    st.markdown("---")
    c_v1, c_v2 = st.columns(2)
    with c_v1:
        st.markdown("#### A. Berat Irisan & Tegangan Overburden (W)")
        st.latex(r"W = b \cdot \left[ \gamma_{\text{timb}} \cdot h_{\text{timb}} + \sum (\gamma_i \cdot h_i) \right] + (q \cdot b)_{x \le B}")
        st.markdown("#### B. Sudut Kemiringan Busur Dasar Irisan (α)")
        st.latex(r"\sin \alpha = \frac{x_c - x_m}{R} \implies \alpha = \arcsin\left(\frac{x_c - x_m}{R}\right)")
    with c_v2:
        st.markdown(r"#### C. Tekanan Air Pori & Rasio Ru ($u \cdot b$)")
        st.latex(r"r_u = \frac{u}{\gamma \cdot z} \implies u \cdot b = r_u \cdot W_{\text{tanah}}")
        st.markdown("#### D. Standar Faktor Keamanan (SNI 8460:2017)")
        st.markdown("• **Statis Jangka Panjang**: FS ≥ 1.50\n• **Jangka Pendek (Akhir Konstruksi)**: FS ≥ 1.30\n• **Beban Gempa (Pseudo-statis)**: FS ≥ 1.10")

# TAB 5: ASAL DATA
with tab5:
    st.markdown('<div class="info-card"><h3>📋 Matriks Pelacakan Asal Data (Data Traceability Matrix)</h3>Menjelaskan dari mana setiap parameter diperoleh dan bagaimana perannya dalam rumus kestabilan lereng.</div>', unsafe_allow_html=True)
    trace_rows = [
        {"Kategori": "🏗️ TIMBUNAN", "Parameter": "H (Tinggi)", "Nilai": f"{P0['H']} m", "Asal Data": "Gambar Geometri Jalan Tol (DED)", "Peran dalam Rumus": "Menentukan profil muka tanah y(x) & tinggi irisan h_timb", "Pengaruh thd FS": "Makin tinggi lereng (H ↑) → Beban pendorong ↑ → FS turun"},
        {"Kategori": "🏗️ TIMBUNAN", "Parameter": "B (Lebar Puncak)", "Nilai": f"{P0['B']} m", "Asal Data": "Cross Section Jalan Tol (Bina Marga)", "Peran dalam Rumus": "Batas sebaran beban lalu lintas q di puncak", "Pengaruh thd FS": "Menentukan jangkauan beban lalu lintas"},
        {"Kategori": "🏗️ TIMBUNAN", "Parameter": "m (Lereng 1:m)", "Nilai": f"1 : {P0['m']}", "Asal Data": "Spesifikasi Desain Lereng Timbunan", "Peran dalam Rumus": "Koordinat kaki lereng (toe = B + m·H)", "Pengaruh thd FS": "Makin landai (m ↑) → Bidang geser lebih aman → FS naik"},
        {"Kategori": "🏗️ TIMBUNAN", "Parameter": "γ timbunan", "Nilai": f"{P0['gE']} kN/m³", "Asal Data": "Uji Pemadatan Lab (Proctor SNI 1742)", "Peran dalam Rumus": "W = γ · h · b pada zona timbunan", "Pengaruh thd FS": "Menambah beban W (mempengaruhi dorong & normal)"},
        {"Kategori": "🏗️ TIMBUNAN", "Parameter": "c' timbunan", "Nilai": f"{P0['cE']} kPa", "Asal Data": "Uji Triaksial CD/CU (SNI 2813)", "Peran dalam Rumus": "Kuat geser kohesi penahan c'·b", "Pengaruh thd FS": "Makin tinggi kohesi (c' ↑) → FS naik linear"},
        {"Kategori": "🏗️ TIMBUNAN", "Parameter": "φ' timbunan", "Nilai": f"{P0['phE']}°", "Asal Data": "Uji Triaksial / Direct Shear Lab", "Peran dalam Rumus": "Komponen friksional (W - ub)·tan φ'", "Pengaruh thd FS": "Penahan utama lereng tanah granular/pilihan (FS ↑)"},
        {"Kategori": "🚗 BEBAN", "Parameter": "q (Lalu Lintas)", "Nilai": f"{P0['q']} kPa", "Asal Data": "Standar Beban Jalan Raya (SNI 1725:2016)", "Peran dalam Rumus": "W += q · b (pada irisan x ≤ B)", "Pengaruh thd FS": "Menambah momen pendorong di puncak lereng (FS ↓)"},
        {"Kategori": "💧 AIR PORI", "Parameter": "ru (Pore Pressure)", "Nilai": f"{P0['ru']}", "Asal Data": "Piezometer / Model Hidrologi SEEP/W", "Peran dalam Rumus": "u·b = ru · W_tanah (mereduksi tegangan efektif)", "Pengaruh thd FS": "Makin jenuh (ru ↑) → Tegangan efektif runtuh → FS anjlok"},
        {"Kategori": "🎯 KRITERIA", "Parameter": "Target FS", "Nilai": f"{target}", "Asal Data": "SNI 8460:2017 Pasal 9 (Stabilitas Lereng)", "Peran dalam Rumus": "Ambang batas evaluasi keamanan lereng", "Pengaruh thd FS": "Tolok ukur keputusan perlu/tidaknya perkuatan lereng"},
        {"Kategori": "📊 KOMPUTASI", "Parameter": "FS Bishop", "Nilai": f"{fs:.3f}", "Asal Data": "Hasil Iterasi Algoritma Bishop 1955", "Peran dalam Rumus": "FS = Σ[(c·b + (W-ub)tanφ)/mα] / Σ(W sin α)", "Pengaruh thd FS": "Angka faktor keamanan desain utama"},
    ]
    for k, (nm, t_, g_, c_, ph_) in enumerate(zip(P["Ln"], P["Lt"], P["Lg"], P["Lc"], P["Lph"])):
        trace_rows.append({"Kategori": f"🌍 TANAH L{k+1}", "Parameter": f"{nm} (t={t_}m)", "Nilai": f"γ={g_}, c={c_}, φ={ph_}°", "Asal Data": "Uji Bor Mesin (SPT ASTM D1586) & Lab UDS", "Peran dalam Rumus": f"Karakteristik tanah dasar pada kedalaman lapisan {k+1}", "Pengaruh thd FS": "Lapisan lunak dasar memicu kelongsoran dalam (deep failure)"})
    st.dataframe(pd.DataFrame(trace_rows), hide_index=True, use_container_width=True, height=520)

st.markdown('<div style="text-align:center;padding:1.5rem;color:#64748b;font-size:0.8rem">TolSlope v2.0 — Analisis Geoteknik Lereng Tol · Metode Fellenius (1927) & Bishop (1955) · SNI 8460:2017</div>', unsafe_allow_html=True)
