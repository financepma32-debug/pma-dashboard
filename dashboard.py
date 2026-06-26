"""
dashboard.py  –  PT Pinus Merah Abadi | FAD
Principal Payment Planning Dashboard

Jalankan: streamlit run dashboard.py
URL      : http://localhost:8501
"""

import json
import sys
import logging
from datetime import datetime, date, timedelta
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ──────────────────────────────────────────────────────────────────────────────
#  PATH  –  edit di sini kalau lokasi berubah
# ──────────────────────────────────────────────────────────────────────────────

DASHBOARD_DATA   = Path(__file__).parent / "data" / "Dashboard_Data.xlsx"
LAST_UPDATE_FILE = Path(__file__).parent / "data" / "last_update.json"
LOGO_PATH        = Path(r"D:\PROJECT FAD\MONITORING PRINSIPLE\LOGO.jpg")
LOG_DIR          = Path(__file__).parent / "logs"

PAYMENT_SCHEDULE = {
    "KSNI":  list(range(7)),   # setiap hari
    "NSI":   [0],              # Senin
    "SIMBA": [1, 3],           # Selasa & Kamis
    "MEIJI": [3],              # Kamis
    "NNI":   [0],              # Senin
}

LUNAS_VALUES   = {"lunas", "lunas "}
SIAP_BAYAR_VAL = {"siap bayar"}

# ──────────────────────────────────────────────────────────────────────────────
#  LOGGING
# ──────────────────────────────────────────────────────────────────────────────

LOG_DIR.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_DIR / "dashboard.log", encoding="utf-8")],
)

# ──────────────────────────────────────────────────────────────────────────────
#  PAGE CONFIG
# ──────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Principal Payment Planning",
    page_icon="🔴",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────────────────────────
#  CSS
# ──────────────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    color: #212529;
}
.main .block-container { padding: 0 1.5rem 2rem 1.5rem; max-width: 1400px; }
#MainMenu, footer, .stDeployButton { display: none !important; }

/* ── Header ── */
.top-header {
    display: flex; align-items: center; justify-content: space-between;
    padding: 0.8rem 1.4rem;
    background: #fff;
    border-bottom: 2px solid #C0392B;
    margin-bottom: 1.25rem;
    box-shadow: 0 1px 4px rgba(0,0,0,.06);
}
.top-header img { height: 36px; object-fit: contain; }
.brand-title   { font-size: .95rem; font-weight: 700; color: #C0392B; line-height: 1.2; }
.brand-sub     { font-size: .68rem; color: #6C757D; }
.badge-ok      { background:#D4EDDA; color:#155724; border:1px solid #C3E6CB;
                 padding:.2rem .6rem; border-radius:10px; font-size:.7rem; font-weight:600; }
.badge-wait    { background:#FFF3CD; color:#856404; border:1px solid #FFEEBA;
                 padding:.2rem .6rem; border-radius:10px; font-size:.7rem; font-weight:600; }

/* ── Section title ── */
.sec-title {
    font-size: .75rem; font-weight: 700; color: #6C757D;
    text-transform: uppercase; letter-spacing: .06em;
    border-left: 3px solid #C0392B;
    padding-left: .5rem;
    margin: 1.25rem 0 .6rem 0;
}

/* ── KPI card ── */
.kpi-card {
    background: #fff; border: 1px solid #DEE2E6;
    border-radius: 7px; padding: .85rem 1rem;
    box-shadow: 0 1px 3px rgba(0,0,0,.04);
}
.kpi-card.hi { border-left: 3px solid #C0392B; }
.kpi-lbl  { font-size:.68rem; font-weight:600; color:#6C757D;
             text-transform:uppercase; letter-spacing:.05em; margin-bottom:.2rem; }
.kpi-val  { font-size:1.35rem; font-weight:700; color:#212529; }
.kpi-val.red { color:#C0392B; }
.kpi-sub  { font-size:.65rem; color:#ADB5BD; margin-top:.15rem; }

/* ── Status pill ── */
.pill-row { display:flex; flex-wrap:wrap; gap:.5rem; margin-bottom:.75rem; }

/* ── Table ── */
.dataframe { font-size:.77rem !important; }

/* ── Schedule table ── */
.sched-tbl { width:100%; border-collapse:collapse; font-size:.79rem; }
.sched-tbl th { background:#F8F9FA; color:#6C757D; font-size:.67rem;
                font-weight:600; text-transform:uppercase; letter-spacing:.04em;
                padding:.45rem .7rem; border-bottom:1px solid #DEE2E6; text-align:left; }
.sched-tbl td { padding:.45rem .7rem; border-bottom:1px solid #DEE2E6; vertical-align:middle; }
.sched-tbl tr:last-child td { border-bottom:none; }
.tag-today { background:#FADBD8; color:#C0392B; padding:.15rem .45rem;
             border-radius:4px; font-size:.67rem; font-weight:700; }
.tag-soon  { background:#FFF3CD; color:#856404; padding:.15rem .45rem;
             border-radius:4px; font-size:.67rem; font-weight:700; }
.tag-later { background:#F8F9FA; color:#6C757D; padding:.15rem .45rem;
             border-radius:4px; font-size:.67rem; }

/* ── Sidebar ── */
[data-testid="stSidebar"] { background:#F8F9FA; border-right:1px solid #DEE2E6; }
.stDownloadButton > button {
    background:#F8F9FA; border:1px solid #DEE2E6; color:#212529;
    font-size:.76rem; border-radius:5px; padding:.35rem .8rem;
}
.stDownloadButton > button:hover { border-color:#C0392B; color:#C0392B; background:#FADBD8; }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
#  HELPER FUNCTIONS
# ──────────────────────────────────────────────────────────────────────────────

def fmt_rp(v, short=True):
    if pd.isna(v) or v == 0: return "Rp 0"
    sign = "-" if v < 0 else ""
    a = abs(v)
    if short:
        if a >= 1e12: return f"{sign}Rp {a/1e12:.2f} T"
        if a >= 1e9:  return f"{sign}Rp {a/1e9:.1f} M"
        if a >= 1e6:  return f"{sign}Rp {a/1e6:.1f} Jt"
    return f"{sign}Rp {a:,.0f}"

def next_schedule(from_date=None):
    if from_date is None:
        from_date = date.today()
    result = {}
    for p, days in PAYMENT_SCHEDULE.items():
        if set(days) == set(range(7)):
            result[p] = from_date; continue
        for i in range(8):
            d = from_date + timedelta(days=i)
            if d.weekday() in days:
                result[p] = d; break
    return result

def load_meta():
    if not LAST_UPDATE_FILE.exists():
        return {}
    try:
        return json.loads(LAST_UPDATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}

@st.cache_data(ttl=300, show_spinner=False)
def load_data():
    if not DASHBOARD_DATA.exists():
        return None
    try:
        df = pd.read_excel(DASHBOARD_DATA, sheet_name="Invoice", engine="openpyxl",
                           dtype={"No Invoice": str, "Principal": str,
                                  "Vendor": str, "PIC": str,
                                  "Status ACC": str, "Kategori": str, "Area": str})
        # Tanggal
        for col in ["Tgl Invoice", "Tgl Jatuh Tempo", "Tgl Payment"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce")
        # Numerik
        for col in ["Nominal Invoice", "Nominal Bayar", "Sisa Tagihan", "Aging (Hari)"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        # Flag
        if "Kategori" in df.columns:
            norm = df["Kategori"].astype(str).str.strip().str.lower()
            df["_lunas"]      = norm.isin(LUNAS_VALUES)
            df["_siap_bayar"] = norm.isin(SIAP_BAYAR_VAL)
        return df
    except Exception as e:
        st.error(f"Gagal baca Dashboard_Data.xlsx: {e}")
        return None

def logo_html():
    if LOGO_PATH.exists():
        import base64
        b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode()
        ext = "jpeg" if LOGO_PATH.suffix.lower() in (".jpg",".jpeg") else LOGO_PATH.suffix.lstrip(".")
        return f'<img src="data:image/{ext};base64,{b64}" alt="logo">'
    return '<span style="font-size:1.4rem;font-weight:800;color:#C0392B;">PMA</span>'

def chart_cfg():
    return {"displayModeBar": False, "responsive": True}

BASE_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter,sans-serif", color="#212529", size=11),
    margin=dict(l=6, r=6, t=32, b=6),
    colorway=["#C0392B","#2980B9","#27AE60","#F39C12","#8E44AD","#16A085"],
)


# ──────────────────────────────────────────────────────────────────────────────
#  HEADER
# ──────────────────────────────────────────────────────────────────────────────

def render_header(meta):
    updated = False
    date_str = meta.get("last_update_date", "—")
    time_str = meta.get("last_update_time", "—")
    try:
        updated = datetime.strptime(date_str, "%Y-%m-%d").date() == date.today()
    except Exception:
        pass
    badge = ('<span class="badge-ok">● Data Updated</span>' if updated
             else '<span class="badge-wait">● Menunggu Update Hari Ini</span>')
    st.markdown(f"""
    <div class="top-header">
        <div style="display:flex;align-items:center;gap:.7rem;">
            {logo_html()}
            <div>
                <div class="brand-title">Principal Payment Planning Dashboard</div>
                <div class="brand-sub">PT Pinus Merah Abadi &nbsp;·&nbsp; Finance Account Payable</div>
            </div>
        </div>
        <div style="text-align:right;">
            <div style="font-size:.68rem;color:#6C757D;margin-bottom:.25rem;">
                Last Update &nbsp;<b style="color:#212529;">{date_str}</b>&nbsp;{time_str}
                &nbsp;|&nbsp; Dashboard_Data.xlsx
            </div>
            {badge}
        </div>
    </div>
    """, unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
#  KPI CARDS
# ──────────────────────────────────────────────────────────────────────────────

def render_kpi(df):
    st.markdown('<div class="sec-title">Executive Summary</div>', unsafe_allow_html=True)

    lunas      = df["_lunas"] if "_lunas" in df.columns else pd.Series(False, index=df.index)
    siap       = df["_siap_bayar"] if "_siap_bayar" in df.columns else pd.Series(False, index=df.index)
    nominal    = df["Nominal Invoice"] if "Nominal Invoice" in df.columns else pd.Series(0.0)

    outstanding = float(nominal[~lunas].sum())
    rtp         = float(nominal[siap].sum())
    avg6        = _avg6(df)
    forecast    = min(avg6, rtp) if avg6 > 0 else rtp
    backlog     = max(rtp - forecast, 0)

    # Baris 1
    c1,c2,c3,c4 = st.columns(4)
    _kpi(c1, "Total Outstanding",  fmt_rp(outstanding), "Invoice belum Lunas", hi=True)
    _kpi(c2, "Siap Bayar",         fmt_rp(rtp),         "Menunggu pembayaran",  hi=True)
    _kpi(c3, "Forecast Payment",   fmt_rp(forecast),    f"Avg 6 bln: {fmt_rp(avg6)}", hi=True)
    _kpi(c4, "Backlog",            fmt_rp(backlog),     "Siap Bayar – Forecast")

    st.markdown("<div style='height:.4rem'></div>", unsafe_allow_html=True)

    # Baris 2
    c5,c6,c7,c8 = st.columns(4)
    _kpi(c5, "Jumlah Invoice",  f"{len(df):,}", "Total baris")
    _kpi(c6, "Jumlah Vendor",   f"{df['Vendor'].nunique() if 'Vendor' in df.columns else 0:,}", "Vendor aktif")
    _kpi(c7, "Principal",       f"{df['Principal'].nunique() if 'Principal' in df.columns else 0:,}", "Principal aktif")
    avg_aging = df["Aging (Hari)"].mean() if "Aging (Hari)" in df.columns else 0
    _kpi(c8, "Rata-rata Aging", f"{avg_aging:.0f} hari", "Rata-rata hari outstanding")

    return {"outstanding": outstanding, "rtp": rtp, "forecast": forecast,
            "backlog": backlog, "avg6": avg6}

def _kpi(col, label, value, sub, hi=False):
    cls = "kpi-card hi" if hi else "kpi-card"
    vcls = "kpi-val red" if hi else "kpi-val"
    col.markdown(f"""
    <div class="{cls}">
        <div class="kpi-lbl">{label}</div>
        <div class="{vcls}">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

def _avg6(df):
    """Rata-rata nominal bayar per bulan, 6 bulan terakhir."""
    lunas = df.get("_lunas", pd.Series(False, index=df.index))
    if "Tgl Payment" in df.columns and "Nominal Bayar" in df.columns:
        src = df[lunas & (df["Nominal Bayar"] > 0)].copy()
        if not src.empty:
            src["_p"] = src["Tgl Payment"].dt.to_period("M")
            monthly = src.groupby("_p")["Nominal Bayar"].sum().sort_index().tail(6)
            if len(monthly): return float(monthly.mean())
    # Fallback via Nominal Invoice
    if "Nominal Invoice" in df.columns and "Bulan" in df.columns and "Tahun" in df.columns:
        src = df[lunas].copy()
        if not src.empty:
            monthly = src.groupby(["Tahun","Bulan"])["Nominal Invoice"].sum().tail(6)
            if len(monthly): return float(monthly.mean())
    return 0.0


# ──────────────────────────────────────────────────────────────────────────────
#  STATUS / KATEGORI PILLS
# ──────────────────────────────────────────────────────────────────────────────

def render_kategori_pills(df):
    st.markdown('<div class="sec-title">Distribusi Kategori</div>', unsafe_allow_html=True)
    if "Kategori" not in df.columns:
        st.info("Kolom Kategori tidak ditemukan.")
        return

    dist = (
        df.groupby("Kategori")
        .agg(Count=("No Invoice","count"),
             Amount=("Nominal Invoice","sum"))
        .reset_index()
        .sort_values("Amount", ascending=False)
    )

    selected = st.session_state.get("sel_kat", None)
    cols = st.columns(min(len(dist), 5))
    for i, row in dist.iterrows():
        col = cols[i % len(cols)]
        is_active = selected == row["Kategori"]
        border = "2px solid #C0392B" if is_active else "1px solid #DEE2E6"
        bg     = "#FADBD8" if is_active else "#F8F9FA"
        if col.button(
            f"{row['Kategori']}\n{int(row['Count'])} invoice\n{fmt_rp(row['Amount'])}",
            key=f"pill_{row['Kategori']}",
            use_container_width=True
        ):
            st.session_state["sel_kat"] = None if is_active else row["Kategori"]
            st.rerun()


# ──────────────────────────────────────────────────────────────────────────────
#  CHARTS
# ──────────────────────────────────────────────────────────────────────────────

def chart_trend(df):
    lunas = df.get("_lunas", pd.Series(False, index=df.index))
    src   = df[lunas].copy()
    if "Tgl Payment" in src.columns and "Nominal Bayar" in src.columns:
        src = src[src["Nominal Bayar"] > 0]
        if not src.empty:
            src["Period"] = src["Tgl Payment"].dt.to_period("M")
            trend = src.groupby("Period")["Nominal Bayar"].sum().sort_index().tail(12).reset_index()
            trend["Period"] = trend["Period"].astype(str)
            trend["MA6"]    = trend["Nominal Bayar"].rolling(6, min_periods=1).mean()

            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=trend["Period"], y=trend["Nominal Bayar"],
                name="Nominal Bayar",
                marker_color="#FADBD8", marker_line_color="#C0392B", marker_line_width=1,
                hovertemplate="<b>%{x}</b><br>%{customdata}<extra></extra>",
                customdata=[fmt_rp(v) for v in trend["Nominal Bayar"]],
            ))
            fig.add_trace(go.Scatter(
                x=trend["Period"], y=trend["MA6"],
                name="Rata-rata 6 Bln", mode="lines+markers",
                line=dict(color="#C0392B", width=2, dash="dash"),
                marker=dict(size=5),
                hovertemplate="<b>%{x}</b><br>Avg: %{customdata}<extra></extra>",
                customdata=[fmt_rp(v) for v in trend["MA6"]],
            ))
            fig.update_layout(
                title=dict(text="Trend Pembayaran Bulanan", font=dict(size=12)),
                xaxis=dict(tickfont=dict(size=10), gridcolor="#DEE2E6"),
                yaxis=dict(tickfont=dict(size=10), gridcolor="#DEE2E6"),
                bargap=0.35, **BASE_LAYOUT
            )
            return fig
    return None

def chart_kategori_bar(df):
    if "Kategori" not in df.columns: return None
    dist = (df.groupby("Kategori")["Nominal Invoice"]
            .sum().sort_values().reset_index())
    colors = ["#C0392B" if "Lunas" in k else
              "#27AE60" if "Siap" in k else "#2980B9"
              for k in dist["Kategori"]]
    fig = go.Figure(go.Bar(
        x=dist["Nominal Invoice"], y=dist["Kategori"],
        orientation="h", marker_color=colors,
        text=[fmt_rp(v) for v in dist["Nominal Invoice"]],
        textposition="outside",
        hovertemplate="<b>%{y}</b><br>%{text}<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text="Nominal per Kategori", font=dict(size=12)),
        xaxis=dict(tickfont=dict(size=9), gridcolor="#DEE2E6"),
        yaxis=dict(tickfont=dict(size=10), gridcolor="rgba(0,0,0,0)"),
        **BASE_LAYOUT
    )
    fig.update_layout(margin=dict(l=10, r=80, t=32, b=6))
    return fig

def chart_aging(df):
    if "Aging Bucket" not in df.columns: return None
    dist = df.groupby("Aging Bucket").agg(
        Count=("No Invoice","count"),
        Amount=("Nominal Invoice","sum")
    ).reset_index()
    bucket_colors = {
        "0-7 Hari": "#27AE60", "8-14 Hari": "#F39C12",
        "15-30 Hari": "#E67E22", "31+ Hari": "#C0392B",
    }
    colors = [bucket_colors.get(b, "#2980B9") for b in dist["Aging Bucket"]]
    fig = go.Figure(go.Bar(
        x=dist["Aging Bucket"], y=dist["Count"],
        marker_color=colors,
        text=dist["Count"].astype(str), textposition="outside",
        hovertemplate="<b>%{x}</b><br>Count: %{y}<br>%{customdata}<extra></extra>",
        customdata=[fmt_rp(v) for v in dist["Amount"]],
    ))
    fig.update_layout(
        title=dict(text="Distribusi Aging", font=dict(size=12)),
        xaxis=dict(tickfont=dict(size=10), gridcolor="rgba(0,0,0,0)"),
        yaxis=dict(tickfont=dict(size=10), gridcolor="#DEE2E6"),
        bargap=0.4, **BASE_LAYOUT
    )
    return fig

def chart_principal(df):
    if "Principal" not in df.columns: return None
    lunas = df.get("_lunas", pd.Series(False, index=df.index))
    grp = (df[~lunas].groupby("Principal")["Nominal Invoice"]
           .sum().sort_values().reset_index())
    fig = go.Figure(go.Bar(
        x=grp["Nominal Invoice"], y=grp["Principal"],
        orientation="h", marker_color="#C0392B", marker_opacity=0.8,
        text=[fmt_rp(v) for v in grp["Nominal Invoice"]], textposition="outside",
        hovertemplate="<b>%{y}</b><br>%{text}<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text="Outstanding per Principal", font=dict(size=12)),
        xaxis=dict(tickfont=dict(size=9), gridcolor="#DEE2E6"),
        yaxis=dict(tickfont=dict(size=10), gridcolor="rgba(0,0,0,0)"),
        **BASE_LAYOUT
    )
    fig.update_layout(margin=dict(l=10, r=80, t=32, b=6))
    return fig

def render_chart(fig, height=290, key=""):
    if fig is None:
        st.markdown(
            f"<div style='height:{height}px;display:flex;align-items:center;"
            "justify-content:center;color:#ADB5BD;font-size:.8rem;"
            "background:#F8F9FA;border:1px dashed #DEE2E6;border-radius:6px;'>"
            "Tidak ada data</div>", unsafe_allow_html=True)
    else:
        st.plotly_chart(fig, use_container_width=True, config=chart_cfg(), key=key or None)


# ──────────────────────────────────────────────────────────────────────────────
#  PAYMENT SCHEDULE
# ──────────────────────────────────────────────────────────────────────────────

# Berapa kali principal bayar per minggu
FREKUENSI_BAYAR = {
    "KSNI":  7,   # setiap hari
    "NSI":   1,   # 1x seminggu (Senin)
    "SIMBA": 2,   # 2x seminggu (Selasa & Kamis)
    "MEIJI": 1,   # 1x seminggu (Kamis)
    "NNI":   1,   # 1x seminggu (Senin)
}

def _forecast_per_principal(df_raw, principal, next_pay_date):
    """
    Forecast payment untuk tanggal next_pay_date (satu sesi bayar).
    Logika:
      - Hitung total bayar 6 bulan terakhir per principal
      - Bagi dengan total sesi bayar dalam 6 bulan sesuai frekuensi
        (mis. KSNI: 26*7=182 hari, SIMBA: 26*2=52 sesi)
      - Hasilnya = rata-rata per SESI bayar
      - Bandingkan dengan Siap Bayar → ambil MIN
    """
    if df_raw is None or df_raw.empty:
        return 0.0

    p_df = df_raw[df_raw["Principal"] == principal] if "Principal" in df_raw.columns else df_raw
    lunas = p_df.get("_lunas", pd.Series(False, index=p_df.index))

    # Frekuensi sesi bayar per minggu
    freq = FREKUENSI_BAYAR.get(principal, 1)
    total_sesi = 26 * freq  # 6 bulan = 26 minggu

    # Total bayar 6 bulan terakhir
    total_paid = 0.0
    if "Tgl Payment" in p_df.columns and "Nominal Bayar" in p_df.columns:
        paid = p_df[lunas & (p_df["Nominal Bayar"] > 0)].copy()
        if not paid.empty:
            cutoff = pd.Timestamp(date.today()) - pd.DateOffset(months=6)
            paid = paid[paid["Tgl Payment"] >= cutoff]
            total_paid = float(paid["Nominal Bayar"].sum())

    # Rata-rata per sesi
    avg_per_sesi = total_paid / total_sesi if total_sesi > 0 else 0.0

    # Siap Bayar s.d. tanggal bayar berikutnya
    siap = 0.0
    siap_mask = p_df.get("_siap_bayar", pd.Series(False, index=p_df.index))
    if "Tgl Jatuh Tempo" in p_df.columns and "Nominal Invoice" in p_df.columns:
        due = pd.Timestamp(next_pay_date)
        siap = float(p_df.loc[siap_mask & (p_df["Tgl Jatuh Tempo"] <= due), "Nominal Invoice"].sum())
        if siap == 0:
            siap = float(p_df.loc[siap_mask, "Nominal Invoice"].sum())

    if avg_per_sesi <= 0 and siap <= 0:
        return 0.0
    if avg_per_sesi <= 0:
        return siap
    return min(avg_per_sesi, siap) if siap > 0 else avg_per_sesi

def render_schedule(df_raw=None):
    st.markdown('<div class="sec-title">Jadwal Pembayaran</div>', unsafe_allow_html=True)
    today    = date.today()
    schedule = next_schedule(today)
    rows     = ""
    for p, d in schedule.items():
        delta = (d - today).days
        if delta == 0:   tag = f'<span class="tag-today">Hari Ini</span>'
        elif delta <= 3: tag = f'<span class="tag-soon">Dalam {delta} Hari</span>'
        else:            tag = f'<span class="tag-later">Dalam {delta} Hari</span>'

        # Forecast minggu ini per principal
        fcst = _forecast_per_principal(df_raw, p, d) if df_raw is not None else 0.0
        fcst_str = fmt_rp(fcst) if fcst > 0 else '<span style="color:#ADB5BD">—</span>'

        rows += (
            f"<tr>"
            f"<td><b>{p}</b></td>"
            f"<td>{d.strftime('%A, %d %B %Y')}</td>"
            f"<td>{tag}</td>"
            f"<td style='font-weight:600;color:#C0392B;'>{fcst_str}</td>"
            f"</tr>"
        )
    st.markdown(f"""
    <table class="sched-tbl">
        <thead>
            <tr>
                <th>Principal</th>
                <th>Tanggal Pembayaran</th>
                <th>Status</th>
                <th>Forecast Payment</th>
            </tr>
        </thead>
        <tbody>{rows}</tbody>
    </table>""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
#  FORECAST BOX
# ──────────────────────────────────────────────────────────────────────────────

def render_forecast_box(kpis):
    st.markdown('<div class="sec-title">Forecast Payment</div>', unsafe_allow_html=True)
    items = [
        ("Siap Bayar",              fmt_rp(kpis["rtp"])),
        ("Rata-rata Bayar 6 Bln",   fmt_rp(kpis["avg6"])),
        ("Forecast Payment",        fmt_rp(kpis["forecast"])),
        ("Backlog",                 fmt_rp(kpis["backlog"])),
    ]
    cols = st.columns(4)
    for col, (label, val) in zip(cols, items):
        col.markdown(f"""
        <div style="background:#F8F9FA;border:1px solid #DEE2E6;border-radius:6px;
                    padding:.7rem .9rem;">
            <div style="font-size:.67rem;font-weight:600;color:#6C757D;
                        text-transform:uppercase;letter-spacing:.05em;">{label}</div>
            <div style="font-size:1.1rem;font-weight:700;color:#C0392B;">{val}</div>
        </div>""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
#  INVOICE TABLE
# ──────────────────────────────────────────────────────────────────────────────

def render_tabel(df):
    st.markdown('<div class="sec-title">Detail Invoice</div>', unsafe_allow_html=True)
    if df.empty:
        st.info("Tidak ada data sesuai filter.")
        return

    SHOW_COLS = [c for c in [
        "Principal", "Vendor", "PIC",
        "No Invoice", "Tgl Invoice", "Tgl Jatuh Tempo",
        "Area", "Nominal Invoice", "Nominal Bayar", "Sisa Tagihan",
        "Status ACC", "Kategori", "Keterangan",
        "Aging (Hari)", "Aging Bucket",
    ] if c in df.columns]

    disp = df[SHOW_COLS].copy()
    for col in ["Tgl Invoice", "Tgl Jatuh Tempo"]:
        if col in disp.columns:
            disp[col] = pd.to_datetime(disp[col], errors="coerce").dt.strftime("%d-%m-%Y")
    for col in ["Nominal Invoice", "Nominal Bayar", "Sisa Tagihan"]:
        if col in disp.columns:
            disp[col] = pd.to_numeric(disp[col], errors="coerce").apply(
                lambda x: f"Rp {x:,.0f}" if pd.notna(x) and x > 0 else "—")
    if "Aging (Hari)" in disp.columns:
        disp["Aging (Hari)"] = disp["Aging (Hari)"].apply(
            lambda x: f"{int(x)} hr" if pd.notna(x) else "—")

    # Pagination
    PAGE_SIZE   = 50
    total       = len(disp)
    total_pages = max(1, (total - 1) // PAGE_SIZE + 1)

    c1, _, c3 = st.columns([2, 3, 2])
    c1.caption(f"{total:,} invoice")
    page = c3.number_input("", min_value=1, max_value=total_pages,
                           value=1, step=1, label_visibility="collapsed")
    c3.caption(f"Hal. {page}/{total_pages}")

    start = (page - 1) * PAGE_SIZE
    st.dataframe(disp.iloc[start:start + PAGE_SIZE],
                 use_container_width=True, hide_index=True)


# ──────────────────────────────────────────────────────────────────────────────
#  DOWNLOAD
# ──────────────────────────────────────────────────────────────────────────────

def render_download(df, kpis):
    st.markdown('<div class="sec-title">Download Laporan</div>', unsafe_allow_html=True)
    st.caption("Semua download mengikuti filter aktif.")

    import io
    from openpyxl.styles import Font, PatternFill, Alignment

    def to_excel(frames_sheets: list):
        """frames_sheets = [(df, sheet_name), ...]"""
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            for frame, name in frames_sheets:
                frame.to_excel(w, sheet_name=name, index=False)
                ws = w.sheets[name]
                hf = Font(bold=True, color="FFFFFF", size=10)
                hfill = PatternFill("solid", fgColor="C0392B")
                for cell in ws[1]:
                    cell.font, cell.fill = hf, hfill
                    cell.alignment = Alignment(horizontal="center")
                ws.freeze_panes = "A2"
        buf.seek(0)
        return buf.read()

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Siapkan data per laporan
    SHOW = [c for c in [
        "Principal","Vendor","PIC","No Invoice","Tgl Invoice","Tgl Jatuh Tempo",
        "Area","Nominal Invoice","Nominal Bayar","Sisa Tagihan",
        "Status ACC","Kategori","Keterangan","Aging (Hari)","Aging Bucket",
    ] if c in df.columns]
    detail_df = df[SHOW].copy()
    for col in ["Tgl Invoice","Tgl Jatuh Tempo"]:
        if col in detail_df.columns:
            detail_df[col] = pd.to_datetime(detail_df[col], errors="coerce").dt.strftime("%d-%m-%Y")

    lunas = df.get("_lunas", pd.Series(False, index=df.index))
    siap  = df.get("_siap_bayar", pd.Series(False, index=df.index))

    # Principal summary
    def _row(x):
        return pd.Series({
            "Total Invoice":      len(x),
            "Lunas":              int(x["_lunas"].sum()),
            "Siap Bayar":         int(x["_siap_bayar"].sum()),
            "Outstanding":        float(x.loc[~x["_lunas"],"Nominal Invoice"].sum()),
            "Nominal Siap Bayar": float(x.loc[x["_siap_bayar"],"Nominal Invoice"].sum()),
        })
    prin_sum = df.groupby("Principal").apply(_row, include_groups=False).reset_index() \
               if "Principal" in df.columns else pd.DataFrame()

    # Kategori summary
    kat_sum = (df.groupby("Kategori")
               .agg(Jumlah=("No Invoice","count"), Total_Nominal=("Nominal Invoice","sum"))
               .reset_index()
               .sort_values("Total_Nominal", ascending=False)
               ) if "Kategori" in df.columns else pd.DataFrame()

    # Forecast
    forecast_df = pd.DataFrame([
        {"Item": "Total Outstanding",    "Nominal (Rp)": kpis["outstanding"]},
        {"Item": "Siap Bayar",           "Nominal (Rp)": kpis["rtp"]},
        {"Item": "Rata-rata Bayar 6Bln", "Nominal (Rp)": kpis["avg6"]},
        {"Item": "Forecast Payment",     "Nominal (Rp)": kpis["forecast"]},
        {"Item": "Backlog",              "Nominal (Rp)": kpis["backlog"]},
    ])

    c1,c2,c3,c4,c5 = st.columns(5)
    with c1:
        st.download_button("📥 Invoice Detail (.xlsx)",
            to_excel([(detail_df, "Invoice Detail")]),
            f"invoice_detail_{ts}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with c2:
        st.download_button("📥 Invoice Detail (.csv)",
            detail_df.to_csv(index=False).encode("utf-8"),
            f"invoice_detail_{ts}.csv", "text/csv",
            use_container_width=True)
    with c3:
        st.download_button("📥 Ringkasan Principal",
            to_excel([(prin_sum, "Principal")]),
            f"principal_{ts}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with c4:
        st.download_button("📥 Ringkasan Kategori",
            to_excel([(kat_sum, "Kategori")]),
            f"kategori_{ts}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with c5:
        st.download_button("📥 Forecast Report",
            to_excel([(forecast_df, "Forecast")]),
            f"forecast_{ts}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)


# ──────────────────────────────────────────────────────────────────────────────
#  SIDEBAR FILTER
# ──────────────────────────────────────────────────────────────────────────────

def render_sidebar(df):
    with st.sidebar:
        st.markdown("**Filter**")

        def opts(col):
            if col not in df.columns: return []
            return sorted(df[col].dropna().astype(str).unique().tolist())

        search    = st.text_input("Cari No Invoice / Keterangan", placeholder="Ketik...")
        principal = st.multiselect("Principal", opts("Principal"))
        pic       = st.multiselect("PIC",       opts("PIC"))
        kategori  = st.multiselect("Kategori",  opts("Kategori"))
        area      = st.multiselect("Area",      opts("Area"))

        st.markdown("---")
        if st.button("Reset Filter", use_container_width=True):
            st.session_state["sel_kat"] = None
            st.rerun()

    return dict(search=search, principal=principal, pic=pic,
                kategori=kategori, area=area)


def apply_filter(df, f):
    if f["principal"]: df = df[df["Principal"].isin(f["principal"])]
    if f["pic"]      : df = df[df["PIC"].isin(f["pic"])]
    if f["kategori"] and "Kategori" in df.columns:
        df = df[df["Kategori"].isin(f["kategori"])]
    if f["area"] and "Area" in df.columns:
        df = df[df["Area"].isin(f["area"])]
    if f["search"]:
        q = f["search"].strip().lower()
        mask = pd.Series(False, index=df.index)
        for col in ["No Invoice","Principal","Keterangan","Area"]:
            if col in df.columns:
                mask |= df[col].astype(str).str.lower().str.contains(q, na=False)
        df = df[mask]
    return df


# ──────────────────────────────────────────────────────────────────────────────
#  MAIN
# ──────────────────────────────────────────────────────────────────────────────

def main():
    meta = load_meta()
    render_header(meta)

    df_raw = load_data()

    # Sidebar
    filters = render_sidebar(df_raw) if df_raw is not None else {}

    if df_raw is None or df_raw.empty:
        st.markdown("""
        <div style="text-align:center;padding:4rem 2rem;background:#F8F9FA;
                    border:1px dashed #DEE2E6;border-radius:8px;margin:2rem 0;">
            <div style="font-size:2.5rem;">📂</div>
            <div style="font-size:1rem;font-weight:600;margin:.5rem 0;">
                Dashboard_Data.xlsx Belum Tersedia</div>
            <div style="font-size:.82rem;color:#6C757D;">
                Jalankan <b>update.exe</b> untuk generate data terbaru.</div>
        </div>""", unsafe_allow_html=True)
        render_schedule(None)
        return

    df = apply_filter(df_raw, filters)

    # ── KPI ──────────────────────────────────────────────────────────────────
    kpis = render_kpi(df)

    # ── Charts ───────────────────────────────────────────────────────────────
    st.markdown('<div class="sec-title">Analitik</div>', unsafe_allow_html=True)
    r1c1, r1c2 = st.columns([3, 2], gap="medium")
    with r1c1: render_chart(chart_trend(df), key="trend")
    with r1c2: render_chart(chart_kategori_bar(df), key="kat_bar")

    r2c1, r2c2 = st.columns(2, gap="medium")
    with r2c1: render_chart(chart_aging(df), key="aging")
    with r2c2: render_chart(chart_principal(df), key="prin")

    # ── Schedule dengan forecast per principal ───────────────────────────────
    render_schedule(df_raw)

    # ── Detail table ─────────────────────────────────────────────────────────
    render_tabel(df)

    # ── Download ─────────────────────────────────────────────────────────────
    render_download(df, kpis)


if __name__ == "__main__":
    main()
