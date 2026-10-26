"""
Dashboard Penjualan Gudeg Sekar Arum 2025 (Streamlit)
Padanan versi web (index.html): tampilan, filter, dan angkanya sama.

Menjalankan:
    pip install -r requirements.txt
    streamlit run app.py

Berkas data Data_Simulasi_Penjualan_Gudeg_Jogja_2025.xlsx harus berada
di folder yang sama dengan app.py.
"""
import math
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import pandas as pd

DATA_FILE = "Data_Simulasi_Penjualan_Gudeg_Jogja_2025.xlsx"
MN = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
MF = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
      "September", "Oktober", "November", "Desember"]
OU = ["Malioboro", "Wijilan", "Kaliurang", "Godean"]
KA = ["Makan di tempat", "Bawa pulang", "Pesan antar online"]
JH = ["Hari kerja", "Akhir pekan", "Libur nasional"]
KC = ["#8b6f55", "#dc9d3f", "#b8442f"]

CSS = """:root{color-scheme:light;--page:#f3ede1;--ivory:#fbf8ef;--line:#e5dac8;--tx:#3a2415;--mu:#7f6b5a;--brown:#4a2a18;--cream:#f0dcb9;--peach:#eaa97b;--red:#b8442f;--amb:#dc9d3f;--tan:#cdbda6;--soft:#ece3d3;--bgr:#8b6f55;box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
html{scroll-padding-top:env(safe-area-inset-top,0px)}*{box-sizing:border-box}
body{margin:0;background:var(--page);color:var(--tx);font:14px/1.45 Inter,system-ui,sans-serif}
main{padding:6px 8px 20px;min-width:0}
.banner{background:#3a2113;border-radius:6px;text-align:center;padding:16px 20px 14px}
.banner h1{margin:0;font:700 34px/1.2 "Playfair Display",Georgia,serif;color:#fff}.banner p{margin:4px 0 0;font-size:12.5px;color:#d9c7ac}
.chips{display:flex;justify-content:flex-end;flex-wrap:wrap;gap:8px;margin:12px 0}.chips span{background:var(--ivory);border:1px solid var(--line);border-radius:14px;padding:3px 12px;font-size:12px;color:var(--mu)}.chips b{color:var(--tx);font-weight:600}
.kpis,.row{display:grid;gap:14px;margin-bottom:14px}.kpis{grid-template-columns:repeat(4,1fr)}.row{grid-template-columns:1.65fr 1fr}.row.r2{grid-template-columns:1.2fr 1fr}.row.r3{grid-template-columns:repeat(3,1fr)}.hint{margin:10px 0 0;padding:9px 14px;background:var(--ivory);border:1px solid var(--line);border-left:4px solid var(--amb);border-radius:8px;font-size:13px;color:var(--tx)}
.kpi{background:var(--ivory);border:1px solid var(--line);border-left:4px solid;border-radius:10px;padding:13px 16px}
.kpi:nth-child(1){border-left-color:var(--red)}.kpi:nth-child(2){border-left-color:var(--amb)}.kpi:nth-child(3){border-left-color:#7a4a2d}.kpi:nth-child(4){border-left-color:#5f7a4a}
.kpi small{display:block;color:var(--mu);font-size:12.5px}.kpi strong{display:block;font-size:26px;line-height:1.25;margin-top:2px}.kpi em{font-style:normal;font-size:12px;color:var(--mu)}
.card{background:var(--ivory);border:1px solid var(--line);border-radius:12px;padding:16px 18px;min-width:0}
.ch{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;margin-bottom:10px}.ch h2{margin:0;font-size:15px;font-weight:700}.ch span{display:block;color:var(--mu);font-size:12px;margin-top:2px}
.seg{display:flex;gap:6px;flex:none}.seg button{border:1px solid var(--line);background:var(--ivory);color:var(--tx);border-radius:7px;padding:4px 11px;font:500 12px Inter,sans-serif;cursor:pointer}.seg button.on{background:var(--brown);border-color:var(--brown);color:#fff}
svg.ch1{display:block;width:100%;height:auto}.ax{fill:var(--mu);font-size:11px}.vl{fill:var(--tx);font-size:12px;font-weight:600}.gl{stroke:var(--line);stroke-dasharray:3 4}
.note{font-size:12.5px;color:var(--mu);margin:8px 0 0}.note b{color:var(--tx)}.empty{color:var(--mu);padding:30px 0;text-align:center}
.hb{display:grid;grid-template-columns:76px 1fr;align-items:center;gap:10px;margin-bottom:10px}.hb>span{font-size:13px}
.trk{height:28px;border-radius:6px;background:var(--soft)}.trk i{display:flex;align-items:center;justify-content:flex-end;height:100%;border-radius:6px;padding-right:9px;font-size:12px;font-weight:700;color:var(--tx)}.trk i.hi{color:#fff}
.trk.o{display:flex;align-items:center}.trk.o i{flex:none;display:block;height:100%;padding:0}.trk.o b{margin-left:7px;font-size:12px;font-weight:700;white-space:nowrap}.lg{display:flex;flex-wrap:wrap;gap:16px;margin:0 0 12px;font-size:12px;color:var(--mu)}.lg i{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:6px}
.sb{display:grid;grid-template-columns:76px 1fr;align-items:center;gap:10px;margin-bottom:10px}.sb>span{font-size:13px}.stk{display:flex;height:30px;border-radius:6px;overflow:hidden}.stk i{display:flex;align-items:center;padding-left:8px;font-size:12px;font-weight:700;color:#fff;border-right:1px solid var(--ivory)}
footer{margin-top:6px;color:var(--mu);font-size:12px;line-height:1.6}
@media(max-width:1000px){.kpis{grid-template-columns:1fr 1fr}.row,.row.r2,.row.r3{grid-template-columns:1fr}main{padding:14px}.banner h1{font-size:26px}}
@media(max-width:560px){.kpis{grid-template-columns:1fr}}"""

SIDE_CSS = """.brand{display:flex;align-items:center;gap:12px;margin-bottom:22px}.brand svg{width:56px;height:60px;flex:none}
.brand h2{margin:0;font:600 26px/1.1 "Playfair Display",Georgia,serif;color:#f4e4c8}.brand p{margin:4px 0 0;font:500 14px "Playfair Display",Georgia,serif;color:var(--peach)}
.dash{display:flex;align-items:center;justify-content:center;gap:12px;background:var(--cream);color:var(--brown);border-radius:14px;padding:14px;font-weight:700;font-size:16px;margin-bottom:26px}
.dash svg{width:22px;height:22px}
[data-testid="stAppViewContainer"],[data-testid="stMain"]{background:#f3ede1}
[data-testid="stHeader"]{background:transparent}
.block-container{padding-top:1.2rem;padding-left:1rem;padding-right:1rem;max-width:100%}
[data-testid="stSidebar"]{background:#4a2a18}
[data-testid="stSidebar"] .brand{margin-bottom:18px}
.dash{display:flex;align-items:center;justify-content:center;gap:12px;background:#f0dcb9;color:#4a2a18;border-radius:14px;padding:14px;font-weight:700;font-size:16px;margin-bottom:22px}
h3.flt{margin:0 0 10px;font-size:13px;letter-spacing:.6px;text-transform:uppercase;color:#c9b08c;font-weight:600}
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p{color:#f1e2c8;font-size:14px}
[data-testid="stSidebar"] [data-baseweb="select"]>div{background:#f4efe6;border-radius:12px;border:0;color:#3a2415}
[data-testid="stSidebar"] .stButton button{background:#ee6f1e;color:#fff;border:0;border-radius:14px;font-weight:700;padding:.6rem 0}
[data-testid="stSidebar"] .stButton button:hover{background:#d9600f;color:#fff}
"""

BRAND_HTML = """
<div class="brand"><svg viewBox="0 0 60 64" fill="none" stroke="#f1dfc3" stroke-width="1.7" stroke-linejoin="round" stroke-linecap="round"><path d="M30 3v8M30 11l-4.5 14h9zM26.5 25 23 41h14l-3.5-16M18 41h24v5H18zM12 46h36v6H12zM7 52h46v6H7z"/></svg><div><h2>Gudeg Jogja</h2><p>Gudeg Sekar Arum</p></div></div>
<div class="dash"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><rect x="3" y="3" width="7" height="7" rx="2"/><rect x="14" y="3" width="7" height="7" rx="2"/><rect x="3" y="14" width="7" height="7" rx="2"/><rect x="14" y="14" width="7" height="7" rx="2"/></svg>Dashboard</div>
<h3 class="flt">Filter</h3>
"""


# ---------------------------------------------------------------- format angka
def nf(n, d=0):
    """Format angka gaya Indonesia: titik ribuan, koma desimal."""
    x = Decimal(repr(float(n))).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)
    s = f"{abs(x):,.{d}f}".replace(",", "§").replace(".", ",").replace("§", ".")
    return ("-" if x < 0 else "") + s


def jt(n):
    return "Rp" + nf(n / 1e9, 2) + " miliar" if n >= 1e9 else "Rp" + nf(n / 1e6, 1) + " juta"


# ------------------------------------------------------------------ data
def load_data(path=None):
    path = Path(path) if path else Path(__file__).parent / DATA_FILE
    d = pd.read_excel(path, sheet_name="Data_Penjualan")
    o = d["Omzet_Bersih_Rp"]
    if not pd.api.types.is_numeric_dtype(o):  # pada berkas .xlsx kolom ini terbaca teks
        d["Omzet_Bersih_Rp"] = pd.to_numeric(
            o.astype(str).str.replace(",", "").str.strip().replace({"-": "0"}))
    d["Tanggal"] = pd.to_datetime(d["Tanggal"])
    d["m"] = d.Tanggal.dt.month - 1
    d["o"] = d.Outlet.map({n: i for i, n in enumerate(OU)})
    d["k"] = d.Kanal_Penjualan.map({n: i for i, n in enumerate(KA)})
    d["j"] = d.Jenis_Hari.map({n: i for i, n in enumerate(JH)})
    d["rw"] = d.Rating_Rata_rata.fillna(0) * d.Jumlah_Ulasan
    return d


def day_counts(d):
    """Jumlah hari per bulan dan jenis hari: {(bulan, jenis): n}."""
    u = d.drop_duplicates("Tanggal")
    return {k: int(v) for k, v in u.groupby(["m", "j"]).size().items()}


def filter_rows(d, o, k, j, a, b):
    m = d.m.between(a, b)
    if o >= 0:
        m &= d.o == o
    if k >= 0:
        m &= d.k == k
    if j >= 0:
        m &= d.j == j
    return d[m]


# ------------------------------------------------------------ potongan HTML
def _step(mx):
    s = mx * 1.04 / 4
    e = 10 ** math.floor(math.log10(s))
    m = s / e
    c = next(x for x in [1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10] if x >= m)
    return c * e * 4


def trend_block(mode, ms, v):
    mxv, mnv = max(v), min(v)
    pk, lw = v.index(mxv), v.index(mnv)
    W, H, L, R, T, B = 640, 270, 56, 22, 28, 30
    iw, ih = W - L - R, H - T - B
    mx = _step(mxv)
    x = (lambda i: L + iw / 2) if len(ms) == 1 else (lambda i: L + i * iw / (len(ms) - 1))
    y = lambda q: T + ih - q / mx * ih
    lb = (lambda q: nf(q / 1e6) + " jt") if mode == "om" else (lambda q: nf(q))
    s = f'<svg class="ch1" viewBox="0 0 {W} {H}" role="img" aria-label="Grafik garis per bulan">'
    for i in range(5):
        q = mx * i / 4
        s += (f'<line class="gl" x1="{L}" x2="{W - R}" y1="{y(q):.2f}" y2="{y(q):.2f}"/>'
              f'<text class="ax" x="{L - 8}" y="{y(q) + 4:.2f}" text-anchor="end">{lb(q)}</text>')
    pts = " ".join(f"{x(i):.2f},{y(q):.2f}" for i, q in enumerate(v))
    s += (f'<polygon points="{x(0):.2f},{y(0):.2f} {pts} {x(len(v) - 1):.2f},{y(0):.2f}" fill="#6b3a22" opacity=".09"/>'
          f'<polyline points="{pts}" fill="none" stroke="#6b3a22" stroke-width="2.5" stroke-linejoin="round"/>')
    for i, q in enumerate(v):
        hi, lo = i == pk, (i == lw and len(v) > 1)
        r = 6.5 if hi else 5 if lo else 3.5
        s += f'<circle cx="{x(i):.2f}" cy="{y(q):.2f}" r="{r}" fill="{"#b8442f" if hi else "#6b3a22"}"/>'
        if (hi or lo) and len(v) > 1:
            anc = "end" if i == len(v) - 1 else "middle"
            s += f'<text class="vl" x="{x(i):.2f}" y="{y(q) - 12:.2f}" text-anchor="{anc}">{lb(q)}</text>'
        s += f'<text class="ax" x="{x(i):.2f}" y="{H - 8}" text-anchor="middle">{MN[ms[i]]}</text>'
    s += "</svg>"
    vf = (lambda q: jt(q)) if mode == "om" else (lambda q: nf(q) + " transaksi")
    if len(ms) > 1:
        note = (f"Bulan tertinggi: <b>{MF[ms[pk]]}</b> ({vf(mxv)}); "
                f"terendah: <b>{MF[ms[lw]]}</b> ({vf(mnv)}).")
    else:
        note = f"Bulan terpilih: <b>{MF[ms[0]]}</b> ({vf(mxv)})."
    title = "Omzet bersih per bulan (Rp Juta)" if mode == "om" else "Jumlah transaksi per bulan"
    return title, s, note


def _bar_o(label, q, scale, hi, fmt):
    col = "var(--red)" if hi else "var(--tan)"
    return (f'<div class="hb"><span>{label}</span><div class="trk o"><i style="width:{q / scale * 78}%;'
            f'background:{col}"></i><b>{fmt(q)}</b></div></div>')


def build_html(d, DN, o=-1, k=-1, j=-1, a=0, b=11):
    rows = filter_rows(d, o, k, j, a, b)
    c = dict(C1="Semua" if o < 0 else OU[o], C2="Semua" if k < 0 else KA[k],
             C3="Semua" if j < 0 else JH[j], C4=MN[a] if a == b else f"{MN[a]}–{MN[b]}")
    empty = '<p class="empty">Tidak ada data untuk kombinasi filter ini.</p>'
    if rows.empty:
        c.update(K1="–", K2="–", K3="–", K4="–", TREND_BLOCKS=empty, OS="", OUTLET=empty,
                 KANAL=empty, HARI=empty, WIS=empty, OLE=empty, RAT=empty, WN="", ON="", RN="")
        return _fill(c)
    om, tx = float(rows.Omzet_Bersih_Rp.sum()), float(rows.Jumlah_Transaksi.sum())
    ul, rs = float(rows.Jumlah_Ulasan.sum()), float(rows.rw.sum())
    c.update(K1=jt(om), K2=nf(tx), K3="Rp" + nf(round(om / tx)), K4=nf(rs / ul, 2) if ul else "–")

    # ---- tren (dua versi, dipilih lewat tombol di halaman)
    ms = list(range(a, b + 1))
    mo = rows.groupby("m").Omzet_Bersih_Rp.sum().reindex(range(12), fill_value=0)
    mt = rows.groupby("m").Jumlah_Transaksi.sum().reindex(range(12), fill_value=0)
    blocks = ""
    for mode, ser in (("om", mo), ("tx", mt)):
        title, svg, note = trend_block(mode, ms, [float(ser[m]) for m in ms])
        hide = "" if mode == "om" else ' style="display:none"'
        blocks += (f'<div data-mode="{mode}"{hide}><div class="ch"><div><h2>{title}</h2>'
                   f'<span>Bulan tertinggi ditandai merah</span></div>__SEG__</div>'
                   f'<div id="trend-{mode}">{svg}</div><p class="note" id="tn-{mode}">{note}</p></div>')
    c["TREND_BLOCKS"] = blocks

    # ---- omzet per outlet
    oo = rows.groupby("o").Omzet_Bersih_Rp.sum().reindex(range(4), fill_value=0)
    od = sorted([(i, float(q)) for i, q in oo.items() if q > 0], key=lambda t: -t[1])
    big = od[0][1] >= 1e9
    un = 1e9 if big else 1e6
    c["OS"] = "Rp " + ("miliar" if big else "juta") + ", diurutkan dari terbesar"
    c["OUTLET"] = "".join(
        f'<div class="hb"><span>{OU[i]}</span><div class="trk"><i class="{"hi" if n == 0 else ""}" '
        f'style="width:{max(q / od[0][1] * 100, 12)}%;background:{"var(--red)" if n == 0 else "var(--tan)"}">'
        f'{nf(q / un, 2 if big else 1)}</i></div></div>' for n, (i, q) in enumerate(od))

    # ---- kanal per outlet
    ok = rows.pivot_table(index="o", columns="k", values="Omzet_Bersih_Rp", aggfunc="sum").reindex(
        index=range(4), columns=range(3)).fillna(0)
    lg = '<div class="lg">' + "".join(
        f'<span><i style="background:{KC[n]}"></i>{nm}</span>' for n, nm in enumerate(KA)) + "</div>"
    kan = ""
    for i, _ in od:
        t = float(ok.loc[i].sum())
        segs = ""
        for kk in range(3):
            q = float(ok.loc[i, kk])
            if q > 0:
                lab = nf(q / t * 100) + "%" if q / t >= .08 else ""
                segs += f'<i style="width:{q / t * 100}%;background:{KC[kk]}">{lab}</i>'
        kan += f'<div class="sb"><span>{OU[i]}</span><div class="stk">{segs}</div></div>'
    c["KANAL"] = lg + kan

    # ---- rata-rata omzet harian menurut jenis hari
    jo = rows.groupby("j").Omzet_Bersih_Rp.sum().reindex(range(3), fill_value=0)
    jd = [sum(DN.get((m, jj), 0) for m in ms) for jj in range(3)]
    av = [(jj, float(jo[jj]) / jd[jj] if jd[jj] else 0.0) for jj in range(3)]
    av = [t for t in av if t[1] > 0 and (j < 0 or j == t[0])]
    am = max(t[1] for t in av)
    h = "".join(
        f'<div class="hb"><span>{JH[jj]}</span><div class="trk"><i class="{"hi" if q == am else ""}" '
        f'style="width:{max(q / am * 100, 12)}%;background:{"var(--red)" if q == am else "var(--tan)"}">'
        f'{nf(q / 1e6, 1)}</i></div></div>' for jj, q in av)
    bs = next((t for t in av if t[0] == 0), None)
    if bs and len(av) > 1:
        h += ('<p class="note">Dibanding hari kerja: ' + ", ".join(
            f'{JH[jj].lower()} <b>{"+" if q >= bs[1] else ""}{nf((q / bs[1] - 1) * 100)}%</b>'
            for jj, q in av if jj > 0) + ".</p>")
    c["HARI"] = h

    # ---- profil outlet: wisatawan, oleh-oleh, rating
    g = rows.groupby("o").agg(t=("Jumlah_Transaksi", "sum"), w=("Transaksi_Wisatawan", "sum"),
                              ol=("Omzet_Oleh_Oleh_Rp", "sum"), gu=("Omzet_Gudeg_Rp", "sum"),
                              rw=("rw", "sum"), u=("Jumlah_Ulasan", "sum")).reindex(range(4), fill_value=0)
    ids = [i for i, _ in od]
    pw = [float(g.w[i] / g.t[i] * 100) if g.t[i] else 0.0 for i in ids]
    po = [float(g.ol[i] / (g.gu[i] + g.ol[i]) * 100) if (g.gu[i] + g.ol[i]) else 0.0 for i in ids]
    pr = [float(g.rw[i] / g.u[i]) if g.u[i] else 0.0 for i in ids]

    def bars(vals, fmt, scale, low):
        mxv, mnv = max(vals), min(vals)
        return "".join(_bar_o(OU[ids[n]], q, scale, len(vals) > 1 and (q == mnv if low else q == mxv), fmt)
                       for n, q in enumerate(vals))

    c["WIS"] = bars(pw, lambda q: nf(q, 1) + "%", 100, False)
    c["OLE"] = bars(po, lambda q: nf(q, 1) + "%", 100, False)
    c["RAT"] = bars(pr, lambda q: nf(q, 2) if q else "–", 5, True)
    multi = len(ids) > 1
    top = lambda vals, fmt: f"Tertinggi: <b>{OU[ids[vals.index(max(vals))]]}</b> ({fmt(max(vals))})." if multi else ""
    c["WN"] = top(pw, lambda q: nf(q, 1) + "%")
    c["ON"] = top(po, lambda q: nf(q, 1) + "%")
    c["RN"] = (f"Terendah ditandai merah: <b>{OU[ids[pr.index(min(pr))]]}</b> ({nf(min(pr), 2)})."
               if multi else "")
    return _fill(c)


SEG = ('<div class="seg"><button class="on" data-m="om">Omzet</button>'
       '<button data-m="tx">Transaksi</button></div>')

PAGE = """<!DOCTYPE html><html lang="id"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Playfair+Display:wght@600;700&display=swap" rel="stylesheet">
<style>__CSS__</style></head><body><main>
<header class="banner"><h1>DASHBOARD PENJUALAN 2025</h1></header>
<p class="hint"><b>Cara membaca:</b> pilih outlet, kanal, jenis hari, atau rentang bulan di panel filter; kartu angka dan semua grafik langsung menyesuaikan. Tekan &ldquo;Reset filter&rdquo; untuk kembali ke seluruh data.</p>
<div class="chips"><span>Outlet: <b id="c1">{{C1}}</b></span><span>Kanal: <b id="c2">{{C2}}</b></span><span>Hari: <b id="c3">{{C3}}</b></span><span>Periode: <b id="c4">{{C4}}</b></span></div>
<section class="kpis">
<div class="kpi"><small>Omzet bersih</small><strong id="k1">{{K1}}</strong></div>
<div class="kpi"><small>Jumlah transaksi</small><strong id="k2">{{K2}}</strong></div>
<div class="kpi"><small>Rata-rata nilai transaksi</small><strong id="k3">{{K3}}</strong></div>
<div class="kpi"><small>Rating berbobot</small><strong id="k4">{{K4}}</strong><em>skala 1–5</em></div>
</section>
<section class="row">
<div class="card" id="trend">{{TREND_BLOCKS}}</div>
<div class="card"><div class="ch"><div><h2>Omzet bersih per outlet</h2><span id="os">{{OS}}</span></div></div><div id="outlet">{{OUTLET}}</div></div>
</section>
<section class="row r2">
<div class="card"><div class="ch"><div><h2>Pangsa kanal penjualan per outlet</h2><span>% omzet bersih</span></div></div><div id="kanal">{{KANAL}}</div></div>
<div class="card"><div class="ch"><div><h2>Rata-rata omzet harian menurut jenis hari</h2><span>Rp juta per hari</span></div></div><div id="hari">{{HARI}}</div></div>
</section>
<section class="row r3">
<div class="card"><div class="ch"><div><h2>Ketergantungan pada wisatawan</h2><span>% transaksi dari wisatawan, per outlet</span></div></div><div id="wis">{{WIS}}</div><p class="note" id="wn">{{WN}}</p></div>
<div class="card"><div class="ch"><div><h2>Peran oleh-oleh</h2><span>% omzet gudeg + oleh-oleh, per outlet</span></div></div><div id="ole">{{OLE}}</div><p class="note" id="on">{{ON}}</p></div>
<div class="card"><div class="ch"><div><h2>Rating pelanggan</h2><span>rating berbobot (skala 1–5), per outlet</span></div></div><div id="rat">{{RAT}}</div><p class="note" id="rn">{{RN}}</p></div>
</section>
<footer>Data: 1 Januari sampai 31 Desember 2025. Sumber: data simulasi penjualan Gudeg Sekar Arum (fiktif), diolah kelompok.<br>Omzet bersih = omzet gudeg + omzet oleh-oleh − diskon. Rating berbobot = Σ(rating × ulasan) ÷ Σ ulasan. Pangsa wisatawan = Σ transaksi wisatawan ÷ Σ transaksi. Pangsa oleh-oleh = Σ omzet oleh-oleh ÷ Σ (omzet gudeg + oleh-oleh). Laba tidak ditampilkan karena data tidak memuat biaya, dan tanggal libur nasional bersifat perkiraan.</footer>
</main>
<script>
function setMode(m){document.querySelectorAll('[data-mode]').forEach(function(e){e.style.display=e.dataset.mode===m?'':'none'});
document.querySelectorAll('.seg button').forEach(function(b){b.classList.toggle('on',b.dataset.m===m)})}
document.querySelectorAll('.seg button').forEach(function(b){b.onclick=function(){setMode(b.dataset.m)}});
</script></body></html>"""


def _fill(c):
    html = PAGE.replace("__CSS__", CSS)
    for key, val in c.items():
        html = html.replace("{{" + key + "}}", val)
    return html.replace("__SEG__", SEG)


# ---------------------------------------------------------------- Streamlit
def main():
    import streamlit as st
    import streamlit.components.v1 as components

    st.set_page_config(page_title="Dashboard Penjualan 2025", page_icon="🍛", layout="wide")
    st.markdown("<style>" + SIDE_CSS + "</style>", unsafe_allow_html=True)

    @st.cache_data
    def get_data():
        d = load_data()
        return d, day_counts(d)

    d, DN = get_data()
    ss = st.session_state
    for key, val in dict(fo=-1, fk=-1, fj=-1, fa=0, fb=11).items():
        ss.setdefault(key, val)

    def fix_a():
        if ss.fa > ss.fb:
            ss.fb = ss.fa

    def fix_b():
        if ss.fb < ss.fa:
            ss.fa = ss.fb

    def reset():
        ss.fo, ss.fk, ss.fj, ss.fa, ss.fb = -1, -1, -1, 0, 11

    with st.sidebar:
        st.markdown(BRAND_HTML, unsafe_allow_html=True)
        st.selectbox("Outlet", [-1, 0, 1, 2, 3], key="fo",
                     format_func=lambda i: "Semua outlet" if i < 0 else OU[i])
        st.selectbox("Kanal penjualan", [-1, 0, 1, 2], key="fk",
                     format_func=lambda i: "Semua kanal" if i < 0 else KA[i])
        st.selectbox("Jenis hari", [-1, 0, 1, 2], key="fj",
                     format_func=lambda i: "Semua jenis hari" if i < 0 else JH[i])
        c1, c2 = st.columns(2)
        c1.selectbox("Bulan dari", list(range(12)), key="fa", format_func=lambda i: MF[i], on_change=fix_a)
        c2.selectbox("Bulan sampai", list(range(12)), key="fb", format_func=lambda i: MF[i], on_change=fix_b)
        st.button("Reset filter", on_click=reset, use_container_width=True)

    html = build_html(d, DN, ss.fo, ss.fk, ss.fj, ss.fa, ss.fb)
    components.html(html, height=1500, scrolling=True)


if __name__ == "__main__":
    main()
