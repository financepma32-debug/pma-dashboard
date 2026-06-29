"""
AR OTC Dashboard — Outstanding MT MTI NKA
PT Pinus Merah Abadi | FAD Team
Palet: Merah PMA #B01C2E + Off-white #FAF7F5 + Abu struktural #2B2B2B / #6B6B6B
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from supabase import create_client, Client
from datetime import datetime

st.set_page_config(
    page_title="AR OTC · MTI NKA · PMA",
    page_icon="🔴",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── DESIGN TOKENS — PMA Brand ────────────────────────────────────────────────
# Merah PMA dari logo: pekat, bukan merah terang
# Off-white "tissue": #FAF7F5 — warm, tidak krem, tidak putih dingin
# Aksent abu gelap untuk teks dan border: terkesan laporan, bukan playful

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

/* ── Reset & base ── */
html, body, [class*="css"]  { font-family: 'Inter', sans-serif; }
[data-testid="stAppViewContainer"] {
    background: #FAF7F5;
}
[data-testid="stSidebar"] {
    background: #FFFFFF;
    border-right: 1px solid #E8E3E0;
}
[data-testid="stSidebar"] * { font-size: 13px; }
.block-container {
    padding: 1.2rem 2rem 2rem;
    max-width: 1400px;
}

/* ── Header strip merah PMA ── */
.pma-header {
    background: #B01C2E;
    border-radius: 10px;
    padding: 16px 24px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 20px;
}
.pma-header-title {
    color: #FFFFFF;
    font-size: 20px;
    font-weight: 700;
    letter-spacing: -0.3px;
    margin: 0;
}
.pma-header-sub {
    color: rgba(255,255,255,0.70);
    font-size: 12px;
    margin: 3px 0 0;
    font-weight: 400;
}
.pma-header-date {
    color: #FFFFFF;
    font-size: 13px;
    font-weight: 600;
    background: rgba(255,255,255,0.15);
    border-radius: 6px;
    padding: 6px 14px;
    font-family: 'DM Mono', monospace;
}

/* ── KPI cards ── */
.kpi-card {
    background: #FFFFFF;
    border-radius: 8px;
    padding: 16px 14px 14px;
    border-left: 4px solid #B01C2E;
    box-shadow: 0 1px 3px rgba(0,0,0,.06);
    height: 88px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
}
.kpi-card.secondary { border-left-color: #D0C8C5; }
.kpi-card.green     { border-left-color: #1A7F4B; }
.kpi-card.warning   { border-left-color: #C4640A; }
.kpi-card.dark      { border-left-color: #2B2B2B; }
.kpi-label {
    font-size: 10px;
    font-weight: 600;
    color: #9A8F8C;
    text-transform: uppercase;
    letter-spacing: .8px;
    margin: 0;
}
.kpi-value {
    font-size: 22px;
    font-weight: 700;
    color: #1A1A1A;
    font-family: 'DM Mono', monospace;
    margin: 0;
    line-height: 1;
}
.kpi-sub {
    font-size: 10px;
    color: #B0A8A5;
    margin: 0;
}

/* ── Bucket aging strip ── */
.bucket-strip {
    display: flex;
    gap: 6px;
    margin: 4px 0 18px;
}
.bucket-cell {
    flex: 1;
    background: #FFFFFF;
    border-radius: 6px;
    padding: 10px 8px;
    text-align: center;
    border-top: 3px solid #ddd;
    box-shadow: 0 1px 3px rgba(0,0,0,.05);
}
.bucket-cell-label {
    font-size: 9px;
    font-weight: 700;
    color: #9A8F8C;
    text-transform: uppercase;
    letter-spacing: .5px;
    margin: 0;
}
.bucket-cell-val {
    font-size: 13px;
    font-weight: 700;
    color: #1A1A1A;
    font-family: 'DM Mono', monospace;
    margin: 4px 0 0;
    line-height: 1;
}

/* ── Section title ── */
.sec-title {
    font-size: 11px;
    font-weight: 700;
    color: #B01C2E;
    text-transform: uppercase;
    letter-spacing: 1.2px;
    margin: 20px 0 8px;
    padding-bottom: 6px;
    border-bottom: 1px solid #E8E3E0;
}

/* ── Alert badge ── */
.badge-red {
    display: inline-block;
    background: #B01C2E;
    color: #fff;
    font-size: 10px;
    font-weight: 700;
    border-radius: 4px;
    padding: 2px 7px;
    margin-left: 6px;
    vertical-align: middle;
}
.badge-green {
    display: inline-block;
    background: #1A7F4B;
    color: #fff;
    font-size: 10px;
    font-weight: 700;
    border-radius: 4px;
    padding: 2px 7px;
    margin-left: 6px;
    vertical-align: middle;
}

/* ── Tabel Streamlit override ── */
[data-testid="stDataFrame"] thead th {
    background: #F2EDEA !important;
    color: #2B2B2B !important;
    font-size: 11px !important;
    font-weight: 700 !important;
    text-transform: uppercase !important;
    letter-spacing: .5px !important;
}
[data-testid="stDataFrame"] tbody td {
    font-size: 12px !important;
    color: #3A3A3A !important;
}

/* ── Last updated bar ── */
.update-bar {
    background: #F2EDEA;
    border-radius: 6px;
    padding: 7px 14px;
    font-size: 11px;
    color: #6B6060;
    margin-bottom: 18px;
    display: flex;
    justify-content: space-between;
}
</style>
""", unsafe_allow_html=True)

# ── KONSTANTA ────────────────────────────────────────────────────────────────
ALL_COLUMNS = [
    "NAMA AREA","NAMA SALES","No Faktur","NAMA TOKO","KATEGORI OVERDUE",
    "REGION","JENIS OUTLET","ASM","RBM","NOMINAL","GROUPING OS",
    "No Faktur SAP","Tanggal Faktur","Tanggal JT","Nilai Faktur","Saldo Awal",
    "TOP","No Faktur Scylla","No Faktur SAP2","Tgl Target Pelunasan System",
    "Tipe Transaksi","Movement","Saldo Akhir",
    "Kronologi","Jenis BA","No BA","BAP","Tgl Konfirmasi","Action Plan","Penyelesaian",
    "BERAPA HARI?","KELOMPOK",
    "CURRENT","1-7 DAYS","8-30 DAYS","31-60 DAYS","61-90 DAYS",
    "91-120 DAYS","121+ DAYS","<2026",
    "KELOMPOK2","OVERDUE","TANGGAL HARI INII","batas 2025","OVERDUE?",
    "ACTUAL PELUNASAN","TARGET PELUNASAN","DUE DATE","Qty Faktur Gantung",
]
NUMERIC_COLS = {
    "NOMINAL","Nilai Faktur","Saldo Awal","Movement","Saldo Akhir",
    "CURRENT","1-7 DAYS","8-30 DAYS","31-60 DAYS","61-90 DAYS",
    "91-120 DAYS","121+ DAYS","<2026",
    "OVERDUE","OVERDUE?","ACTUAL PELUNASAN","TARGET PELUNASAN",
    "DUE DATE","Qty Faktur Gantung",
}
DATE_COLS = {
    "Tanggal Faktur","Tanggal JT","Tgl Target Pelunasan System",
    "Tgl Konfirmasi","TANGGAL HARI INII","batas 2025",
}
BUCKETS = ["CURRENT","1-7 DAYS","8-30 DAYS","31-60 DAYS",
           "61-90 DAYS","91-120 DAYS","121+ DAYS","<2026"]

# Warna bucket — gradasi dari hijau ke merah PMA, diakhiri abu untuk <2026
BUCKET_BORDER = {
    "CURRENT":    "#1A7F4B",
    "1-7 DAYS":   "#3DAB6E",
    "8-30 DAYS":  "#C4640A",
    "31-60 DAYS": "#D4870A",
    "61-90 DAYS": "#B01C2E",
    "91-120 DAYS":"#8A0E1E",
    "121+ DAYS":  "#5C0813",
    "<2026":      "#6B6B6B",
}

# ── FORMAT ───────────────────────────────────────────────────────────────────
def M(v):
    """Format angka besar ke Miliar/Juta/plain"""
    v = float(v)
    if abs(v) >= 1_000_000_000: return f"{v/1_000_000_000:.2f}M"
    if abs(v) >= 1_000_000:     return f"{v/1_000_000:.1f}Jt"
    return f"{v:,.0f}"

def R(v): return f"{float(v):,.0f}"
def P(n, d): return f"{n/d*100:.1f}%" if d else "–"

# ── SUPABASE ─────────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def get_sb():
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

@st.cache_data(ttl=300, show_spinner=False)
def load_data():
    sb = get_sb()
    last_updated = "–"
    try:
        meta = sb.table("upload_log").select("uploaded_at,total_rows") \
                 .order("uploaded_at", desc=True).limit(1).execute()
        if meta.data:
            raw = meta.data[0].get("uploaded_at","")
            try:
                dt = datetime.fromisoformat(raw.replace("Z","+00:00"))
                last_updated = dt.strftime("%d %b %Y · %H:%M WIB")
            except Exception:
                last_updated = raw
    except Exception:
        pass

    try:
        rows, page, SZ = [], 0, 1000
        while True:
            r = sb.table("os_master").select("*").range(page*SZ,(page+1)*SZ-1).execute()
            if not r.data: break
            rows.extend(r.data)
            if len(r.data) < SZ: break
            page += 1
        if not rows:
            return pd.DataFrame(), last_updated

        df = pd.DataFrame(rows)
        for c in ("id","created_at"):
            if c in df.columns: df.drop(columns=c, inplace=True)
        for col in ALL_COLUMNS:
            if col not in df.columns: df[col] = None
        df = df[ALL_COLUMNS].copy()
        for col in NUMERIC_COLS:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        for col in DATE_COLS:
            df[col] = pd.to_datetime(df[col], errors="coerce")
        return df, last_updated
    except Exception as e:
        st.error(f"Gagal ambil data Supabase: {e}")
        return pd.DataFrame(), last_updated

# ── PLOTLY TEMPLATE PMA ──────────────────────────────────────────────────────
PMA_RED   = "#B01C2E"
PMA_LIGHT = "#FAF7F5"
PLOT_FONT = dict(family="Inter, sans-serif", size=11, color="#2B2B2B")

def pma_layout(fig, height=300, margin=None):
    m = margin or dict(t=16, b=12, l=8, r=8)
    fig.update_layout(
        plot_bgcolor=PMA_LIGHT,
        paper_bgcolor="rgba(0,0,0,0)",
        font=PLOT_FONT,
        height=height,
        margin=m,
        xaxis=dict(gridcolor="#EDE8E6", linecolor="#D0C8C5", tickfont_size=10),
        yaxis=dict(gridcolor="#EDE8E6", linecolor="#D0C8C5", tickfont_size=10),
        showlegend=False,
    )
    return fig

def sec(title, badge=None, badge_type="red"):
    badge_html = ""
    if badge:
        cls = "badge-red" if badge_type=="red" else "badge-green"
        badge_html = f"<span class='{cls}'>{badge}</span>"
    st.markdown(f"<p class='sec-title'>{title}{badge_html}</p>",
                unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════
def main():
    with st.spinner("Mengambil data..."):
        df, last_updated = load_data()

    # ── HEADER ───────────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="pma-header">
      <div>
        <p class="pma-header-title">AR Outstanding MT — MTI NKA</p>
        <p class="pma-header-sub">PT Pinus Merah Abadi · FAD Team · Data via Supabase</p>
      </div>
      <span class="pma-header-date">{datetime.now().strftime('%d %b %Y')}</span>
    </div>
    """, unsafe_allow_html=True)

    if df.empty:
        st.warning("Belum ada data. Jalankan **JALANKAN_UPLOAD.bat** untuk upload data.")
        return

    # ── SIDEBAR ───────────────────────────────────────────────────────────
    st.sidebar.markdown("### Filter", unsafe_allow_html=False)

    def sb(label, col, src):
        opts = ["Semua"] + sorted(src[col].dropna().unique().tolist())
        return st.sidebar.selectbox(label, opts, key=f"f_{col}")

    sel_region = sb("Region",       "REGION",       df)
    d0 = df if sel_region=="Semua" else df[df["REGION"]==sel_region]
    sel_area   = sb("Nama Area",    "NAMA AREA",    d0)
    d1 = d0 if sel_area=="Semua" else d0[d0["NAMA AREA"]==sel_area]
    sel_jenis  = sb("Jenis Outlet", "JENIS OUTLET", d1)
    sel_asm    = sb("ASM",          "ASM",          d1)
    sel_rbm    = sb("RBM",          "RBM",          d1)
    sel_grp    = sb("Grouping OS",  "GROUPING OS",  d1)
    sel_bkt    = st.sidebar.multiselect("Kelompok Aging", BUCKETS, default=BUCKETS)
    st.sidebar.markdown("---")
    if st.sidebar.button("↺  Refresh Data", use_container_width=True):
        st.cache_data.clear(); st.rerun()
    st.sidebar.caption(f"Update: {last_updated}")

    # ── APPLY FILTER ──────────────────────────────────────────────────────
    dff = df.copy()
    if sel_region!="Semua": dff = dff[dff["REGION"]     ==sel_region]
    if sel_area  !="Semua": dff = dff[dff["NAMA AREA"]  ==sel_area]
    if sel_jenis !="Semua": dff = dff[dff["JENIS OUTLET"]==sel_jenis]
    if sel_asm   !="Semua": dff = dff[dff["ASM"]        ==sel_asm]
    if sel_rbm   !="Semua": dff = dff[dff["RBM"]        ==sel_rbm]
    if sel_grp   !="Semua": dff = dff[dff["GROUPING OS"] ==sel_grp]
    if sel_bkt:              dff = dff[dff["KELOMPOK"].isin(sel_bkt)]

    if dff.empty:
        st.warning("Tidak ada data sesuai filter."); return

    # ── UPDATE BAR ────────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="update-bar">
      <span>⏱ Update terakhir: <strong>{last_updated}</strong></span>
      <span>{len(dff):,} faktur ditampilkan</span>
    </div>""", unsafe_allow_html=True)

    # ════════════════════════════════════════════════════════════════════
    # BARIS 1 — 6 KPI UTAMA
    # ════════════════════════════════════════════════════════════════════
    total_nom  = dff["NOMINAL"].sum()
    total_nf   = dff["Nilai Faktur"].sum()
    total_ov   = dff["OVERDUE"].sum()
    total_cur  = dff["CURRENT"].sum()
    total_akt  = dff["ACTUAL PELUNASAN"].sum()
    total_tgt  = dff["TARGET PELUNASAN"].sum()
    total_due  = dff["DUE DATE"].sum()
    total_qty  = int(dff["Qty Faktur Gantung"].sum())
    pct_ov     = total_ov / total_nom * 100 if total_nom else 0
    pct_coll   = total_akt / total_tgt * 100 if total_tgt else 0

    def kpi(col_obj, label, val, sub="", cls=""):
        with col_obj:
            st.markdown(
                f"<div class='kpi-card {cls}'>"
                f"<p class='kpi-label'>{label}</p>"
                f"<p class='kpi-value'>{val}</p>"
                f"<p class='kpi-sub'>{sub}</p>"
                f"</div>", unsafe_allow_html=True)

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    kpi(c1, "Total Outstanding",   M(total_nom), f"Nilai Faktur {M(total_nf)}")
    kpi(c2, "Overdue",             M(total_ov),  f"{pct_ov:.1f}% dari outstanding", "")
    kpi(c3, "Current",             M(total_cur), f"{P(total_cur,total_nom)} dari total", "green")
    kpi(c4, "% Collection",        f"{pct_coll:.1f}%",
        f"Actual {M(total_akt)} / Target {M(total_tgt)}", "warning")
    kpi(c5, "Due Date Hari Ini",   M(total_due), "Nominal jatuh tempo hari ini", "secondary")
    kpi(c6, "Qty Faktur Gantung",  f"{total_qty:,}", f"{len(dff):,} baris data", "dark")

    st.markdown("<br>", unsafe_allow_html=True)

    # ════════════════════════════════════════════════════════════════════
    # BARIS 2 — BUCKET AGING (strip horizontal penuh)
    # ════════════════════════════════════════════════════════════════════
    bv = {b: dff[b].sum() for b in BUCKETS}
    grand = sum(bv.values())

    cells_html = "".join([
        f"<div class='bucket-cell' style='border-top-color:{BUCKET_BORDER[b]}'>"
        f"<p class='bucket-cell-label'>{b}</p>"
        f"<p class='bucket-cell-val'>{M(bv[b])}</p>"
        f"</div>"
        for b in BUCKETS
    ])
    cells_html += (
        f"<div class='bucket-cell' style='border-top-color:#2B2B2B;background:#2B2B2B'>"
        f"<p class='bucket-cell-label' style='color:#FAF7F5'>TOTAL</p>"
        f"<p class='bucket-cell-val' style='color:#FFFFFF'>{M(grand)}</p>"
        f"</div>"
    )
    st.markdown(f"<div class='bucket-strip'>{cells_html}</div>", unsafe_allow_html=True)

    # ════════════════════════════════════════════════════════════════════
    # BARIS 3 — Chart KIRI: Grouping OS | KANAN: Distribusi Jenis Outlet
    # ════════════════════════════════════════════════════════════════════
    sec("Komposisi Outstanding")
    col_a, col_b = st.columns([5,4])

    with col_a:
        grp_os = (dff.groupby("GROUPING OS")["NOMINAL"]
                  .sum().sort_values().reset_index())
        # Buang yang 0 dan negatif kecil (RETUR, LUNAS)
        grp_os = grp_os[grp_os["NOMINAL"] > 0]
        grp_os.columns = ["Kategori","Nominal"]
        # Warna: satu warna solid PMA merah, luminosity bervariasi by rank
        n = len(grp_os)
        colors = [f"rgba(176,28,46,{0.35 + 0.65*(i/max(n-1,1)):.2f})" for i in range(n)]
        fig_os = go.Figure(go.Bar(
            x=grp_os["Nominal"], y=grp_os["Kategori"],
            orientation="h",
            marker_color=colors,
            text=[M(v) for v in grp_os["Nominal"]],
            textposition="outside",
            textfont=dict(size=10, color="#2B2B2B"),
        ))
        pma_layout(fig_os, height=310)
        fig_os.update_layout(xaxis_tickformat=",")
        st.plotly_chart(fig_os, use_container_width=True)

    with col_b:
        # Donut: CURRENT vs tiap bucket overdue
        ov_breakdown = [(b, bv[b]) for b in BUCKETS if b != "CURRENT" and bv[b] > 0]
        pie_labels = ["CURRENT"] + [x[0] for x in ov_breakdown]
        pie_vals   = [bv["CURRENT"]] + [x[1] for x in ov_breakdown]
        pie_colors = (
            [BUCKET_BORDER["CURRENT"]] +
            [BUCKET_BORDER[x[0]] for x in ov_breakdown]
        )
        fig_pie = go.Figure(go.Pie(
            labels=pie_labels, values=pie_vals,
            marker_colors=pie_colors,
            hole=0.55,
            textinfo="percent",
            textfont_size=10,
            hovertemplate="%{label}: %{value:,.0f}<extra></extra>",
        ))
        fig_pie.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            height=310,
            margin=dict(t=16,b=12,l=0,r=0),
            legend=dict(font_size=10, orientation="v",
                        x=1.02, y=0.5, xanchor="left"),
            showlegend=True,
            annotations=[dict(
                text=f"<b>{M(grand)}</b><br><span style='font-size:9px'>Total OS</span>",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=13, color="#2B2B2B"),
            )],
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    # ════════════════════════════════════════════════════════════════════
    # BARIS 4 — TABEL NAMA AREA (Current | Overdue | %Curr/OD)
    #         + TABEL OVERDUE per KATEGORI OVERDUE
    # ════════════════════════════════════════════════════════════════════
    sec("Outstanding per Wilayah & Kategori")
    col_c, col_d = st.columns(2)

    with col_c:
        st.caption("**Per Nama Area** — Current · Overdue · % Curr/OD")
        t_area = (
            dff.groupby("NAMA AREA")
            .agg(
                Nilai_Faktur=("Nilai Faktur","sum"),
                Current=("CURRENT","sum"),
                Overdue=("OVERDUE","sum"),
            )
            .reset_index()
            .sort_values("Nilai_Faktur", ascending=False)
        )
        t_area["% Curr/OD"] = t_area.apply(
            lambda r: P(r["Current"], r["Current"]+r["Overdue"]), axis=1)
        t_area["Nilai Faktur"] = t_area["Nilai_Faktur"].apply(R)
        t_area["Current"]      = t_area["Current"].apply(R)
        t_area["Overdue"]      = t_area["Overdue"].apply(R)
        t_area = t_area[["NAMA AREA","Nilai Faktur","Current","Overdue","% Curr/OD"]]
        t_area.columns = ["Nama Area","Nilai Faktur","Current","Overdue","% Curr/OD"]
        st.dataframe(t_area, use_container_width=True, hide_index=True, height=340)

    with col_d:
        st.caption("**Per Kategori Overdue** (KATEGORI OVERDUE)")
        t_kat = (
            dff[dff["NOMINAL"]>0]
            .groupby("KATEGORI OVERDUE")
            .agg(
                Nominal=("NOMINAL","sum"),
                Faktur=("No Faktur","count"),
            )
            .reset_index()
            .sort_values("Nominal", ascending=False)
        )
        t_kat["% Total"] = t_kat["Nominal"].apply(lambda v: P(v, total_nom))
        t_kat["Nominal"] = t_kat["Nominal"].apply(R)
        t_kat.columns = ["Kategori Overdue","Nominal","Jml Faktur","% Total"]
        st.dataframe(t_kat, use_container_width=True, hide_index=True, height=340)

    # ════════════════════════════════════════════════════════════════════
    # BARIS 5 — COLLECTION: Actual vs Target
    #           Chart grouped bar + tabel ASM collection rate
    # ════════════════════════════════════════════════════════════════════
    sec("Collection — Actual vs Target Pelunasan")
    col_e, col_f = st.columns([5,4])

    with col_e:
        # Grouped bar: top 10 Nama Area, dua bar (Actual & Target)
        t_coll = (
            dff.groupby("NAMA AREA")
            .agg(Actual=("ACTUAL PELUNASAN","sum"),
                 Target=("TARGET PELUNASAN","sum"))
            .reset_index()
        )
        t_coll = t_coll[t_coll["Target"]>0].nlargest(10,"Target")
        fig_coll = go.Figure()
        fig_coll.add_trace(go.Bar(
            name="Target", x=t_coll["NAMA AREA"], y=t_coll["Target"],
            marker_color="#E8BCBF", text=[M(v) for v in t_coll["Target"]],
            textposition="outside", textfont_size=9,
        ))
        fig_coll.add_trace(go.Bar(
            name="Actual", x=t_coll["NAMA AREA"], y=t_coll["Actual"],
            marker_color=PMA_RED, text=[M(v) for v in t_coll["Actual"]],
            textposition="outside", textfont_size=9,
        ))
        pma_layout(fig_coll, height=320)
        fig_coll.update_layout(
            barmode="group",
            showlegend=True,
            legend=dict(orientation="h", x=0, y=1.08, font_size=10),
            xaxis_tickangle=-30,
            yaxis_tickformat=",",
        )
        st.plotly_chart(fig_coll, use_container_width=True)

    with col_f:
        st.caption("**Collection Rate per ASM**")
        t_asm_coll = (
            dff.groupby("ASM")
            .agg(
                Actual=("ACTUAL PELUNASAN","sum"),
                Target=("TARGET PELUNASAN","sum"),
                Outstanding=("NOMINAL","sum"),
            )
            .reset_index()
        )
        t_asm_coll = t_asm_coll[t_asm_coll["Target"]>0].sort_values("Target",ascending=False)
        t_asm_coll["%Coll"] = t_asm_coll.apply(lambda r: P(r["Actual"],r["Target"]), axis=1)
        t_asm_coll["Actual"]      = t_asm_coll["Actual"].apply(M)
        t_asm_coll["Target"]      = t_asm_coll["Target"].apply(M)
        t_asm_coll["Outstanding"] = t_asm_coll["Outstanding"].apply(M)
        t_asm_coll.columns = ["ASM","Actual","Target","Outstanding","% Coll"]
        st.dataframe(t_asm_coll, use_container_width=True, hide_index=True, height=320)

    # ════════════════════════════════════════════════════════════════════
    # BARIS 6 — OVERDUE per JENIS OUTLET + RBM Summary
    # ════════════════════════════════════════════════════════════════════
    sec("Breakdown Jenis Outlet & RBM")
    col_g, col_h = st.columns([4,5])

    with col_g:
        t_outlet = (
            dff.groupby("JENIS OUTLET")
            .agg(Outstanding=("NOMINAL","sum"), Overdue=("OVERDUE","sum"))
            .reset_index()
            .sort_values("Outstanding", ascending=False)
            .head(12)
        )
        t_outlet["% OD"] = t_outlet.apply(lambda r: P(r["Overdue"],r["Outstanding"]), axis=1)
        t_outlet["Outstanding"] = t_outlet["Outstanding"].apply(M)
        t_outlet["Overdue"]     = t_outlet["Overdue"].apply(M)
        t_outlet.columns = ["Jenis Outlet","Outstanding","Overdue","% OD"]
        st.dataframe(t_outlet, use_container_width=True, hide_index=True, height=320)

    with col_h:
        t_rbm = (
            dff.groupby("RBM")
            .agg(
                Outstanding=("NOMINAL","sum"),
                Overdue=("OVERDUE","sum"),
                Actual=("ACTUAL PELUNASAN","sum"),
                Target=("TARGET PELUNASAN","sum"),
            )
            .reset_index()
            .sort_values("Outstanding", ascending=False)
        )
        t_rbm["%OD"]   = t_rbm.apply(lambda r: P(r["Overdue"],r["Outstanding"]), axis=1)
        t_rbm["%Coll"] = t_rbm.apply(lambda r: P(r["Actual"],r["Target"]), axis=1)
        for c in ["Outstanding","Overdue","Actual","Target"]:
            t_rbm[c] = t_rbm[c].apply(M)
        t_rbm.columns = ["RBM","Outstanding","Overdue","Actual Pel.","Target Pel.","%OD","%Coll"]
        st.dataframe(t_rbm, use_container_width=True, hide_index=True, height=320)

    # ════════════════════════════════════════════════════════════════════
    # BARIS 7 — CHART: Top 10 NAMA AREA outstanding
    #         + Distribusi OVERDUE? (hari overdue)
    # ════════════════════════════════════════════════════════════════════
    sec("Top Wilayah & Profil Hari Overdue")
    col_i, col_j = st.columns(2)

    with col_i:
        top_area = (
            dff.groupby("NAMA AREA")["NOMINAL"]
            .sum().nlargest(10).reset_index()
        )
        top_area.columns = ["Area","Nominal"]
        # Warna: area terbesar = merah PMA penuh, sisanya fade
        n = len(top_area)
        bar_colors = [PMA_RED if i==0
                      else f"rgba(176,28,46,{max(0.25, 1-i*0.09):.2f})"
                      for i in range(n)]
        fig_area = go.Figure(go.Bar(
            x=top_area["Nominal"], y=top_area["Area"],
            orientation="h",
            marker_color=bar_colors,
            text=[M(v) for v in top_area["Nominal"]],
            textposition="outside",
            textfont=dict(size=10),
        ))
        pma_layout(fig_area, height=320)
        fig_area.update_layout(yaxis=dict(autorange="reversed"), xaxis_tickformat=",")
        st.plotly_chart(fig_area, use_container_width=True)

    with col_j:
        # Histogram distribusi hari overdue — insight yang tidak ada di dashboard lama
        df_ov = dff[dff["OVERDUE?"]>0].copy()
        if not df_ov.empty:
            bins   = [0,7,30,60,90,120,9999]
            labels = ["1–7 hr","8–30 hr","31–60 hr","61–90 hr","91–120 hr","121+ hr"]
            df_ov["bucket_hari"] = pd.cut(df_ov["OVERDUE?"], bins=bins,
                                          labels=labels, right=True)
            dist = (df_ov.groupby("bucket_hari",observed=True)
                    .agg(Nominal=("NOMINAL","sum"), Faktur=("No Faktur","count"))
                    .reset_index())
            dist.columns = ["Bucket","Nominal","Faktur"]
            pal = ["#3DAB6E","#C4640A","#D4870A","#B01C2E","#8A0E1E","#5C0813"]
            fig_dist = go.Figure(go.Bar(
                x=dist["Bucket"], y=dist["Nominal"],
                marker_color=pal[:len(dist)],
                text=[f"{M(v)}\n({r} fktr)" for v,r in zip(dist["Nominal"],dist["Faktur"])],
                textposition="outside", textfont_size=9,
            ))
            pma_layout(fig_dist, height=320)
            fig_dist.update_layout(yaxis_tickformat=",")
            st.plotly_chart(fig_dist, use_container_width=True)

    # ════════════════════════════════════════════════════════════════════
    # BARIS 8 — DETAIL FAKTUR (expandable)
    # ════════════════════════════════════════════════════════════════════
    ovrd_ct = int((dff["OVERDUE?"]>0).sum())
    sec("Detail Faktur", badge=f"{ovrd_ct:,} faktur overdue", badge_type="red")

    COLS_DET = ["NAMA AREA","RBM","ASM","NAMA SALES","NAMA TOKO",
                "No Faktur","Tanggal Faktur","Tanggal JT",
                "Nilai Faktur","Saldo Akhir","KELOMPOK","OVERDUE?","GROUPING OS"]
    cols_ok = [c for c in COLS_DET if c in dff.columns]
    tbl = dff[cols_ok].copy()
    if "Tanggal Faktur" in tbl.columns:
        tbl["Tanggal Faktur"] = tbl["Tanggal Faktur"].dt.strftime("%d %b %Y")
    if "Tanggal JT" in tbl.columns:
        tbl["Tanggal JT"] = tbl["Tanggal JT"].dt.strftime("%d %b %Y")
    for c in ["Nilai Faktur","Saldo Akhir"]:
        if c in tbl.columns: tbl[c] = tbl[c].apply(R)
    tbl.rename(columns={
        "NAMA AREA":"Nama Area","NAMA SALES":"Nama Sales","NAMA TOKO":"Nama Toko",
        "Saldo Akhir":"Sisa AR","OVERDUE?":"Hari OD","KELOMPOK":"Kelompok",
        "GROUPING OS":"Grouping OS",
    }, inplace=True)
    tbl.insert(0,"#", range(1, len(tbl)+1))

    with st.expander(f"🔽 Tampilkan {len(tbl):,} baris · Nilai Faktur Total: {M(dff['Nilai Faktur'].sum())} · OS: {M(total_nom)}",
                     expanded=False):
        st.dataframe(tbl, use_container_width=True, hide_index=True, height=460)
        st.download_button(
            "⬇️ Download CSV (hasil filter)",
            data=dff[cols_ok].to_csv(index=False,sep=";",encoding="utf-8-sig").encode("utf-8-sig"),
            file_name=f"OS_MTI_NKA_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
        )

    # ── FOOTER ───────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown(
        f"<p style='text-align:center;color:#C4B8B5;font-size:11px;"
        f"font-family:DM Mono,monospace'>"
        f"AR OTC · PT Pinus Merah Abadi · FAD Team · "
        f"data {last_updated} · render {datetime.now().strftime('%d %b %Y %H:%M')}"
        f"</p>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
