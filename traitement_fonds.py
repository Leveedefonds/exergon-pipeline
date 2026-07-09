import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta

st.set_page_config(
    page_title="Reporting Exergon",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Global font */
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    /* Header */
    .main-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        padding: 2rem 2.5rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        color: white;
    }
    .main-header h1 { font-size: 2rem; font-weight: 700; margin: 0; letter-spacing: -0.5px; }
    .main-header p  { margin: 0.25rem 0 0; opacity: 0.65; font-size: 0.9rem; }

    /* KPI cards */
    .kpi-card {
        background: white;
        border-radius: 14px;
        padding: 1.4rem 1.6rem;
        border: 1px solid #eef0f5;
        box-shadow: 0 2px 12px rgba(0,0,0,0.06);
        transition: transform 0.15s;
    }
    .kpi-card:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.1); }
    .kpi-label  { font-size: 0.78rem; font-weight: 600; color: #8b93a7; text-transform: uppercase; letter-spacing: 0.6px; margin-bottom: 0.4rem; }
    .kpi-value  { font-size: 1.9rem; font-weight: 800; color: #1a1a2e; line-height: 1; }
    .kpi-sub    { font-size: 0.8rem; color: #8b93a7; margin-top: 0.4rem; }
    .kpi-badge  { display: inline-block; font-size: 0.72rem; font-weight: 600;
                  padding: 0.15rem 0.55rem; border-radius: 20px; margin-top: 0.4rem; }
    .badge-green  { background:#ecfdf5; color:#059669; }
    .badge-orange { background:#fff7ed; color:#d97706; }
    .badge-red    { background:#fef2f2; color:#dc2626; }
    .badge-blue   { background:#eff6ff; color:#2563eb; }

    /* Section headers */
    .section-title {
        font-size: 1.05rem; font-weight: 700; color: #1a1a2e;
        margin: 1.8rem 0 1rem; padding-bottom: 0.5rem;
        border-bottom: 2px solid #e8eaf0;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: #f8f9fc;
        border-right: 1px solid #e8eaf0;
    }
    [data-testid="stSidebar"] .sidebar-logo {
        font-size: 1.2rem; font-weight: 800; color: #0f3460;
        padding: 0.5rem 0 1.5rem;
    }

    /* Tab overrides */
    [data-testid="stTabs"] button {
        font-weight: 600; font-size: 0.9rem;
    }

    /* Dataframe */
    [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

    /* Hide streamlit branding */
    #MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─── DATA ───────────────────────────────────────────────────────────────────────
@st.cache_data
def load_data(path):
    df = pd.read_excel(path, header=2)

    # Revenu en M€ (numérique)
    df["Revenu_M"] = pd.to_numeric(df["Revenu attendu"], errors="coerce").fillna(0) / 1_000_000

    # Maturation numérique
    df["Matu_num"] = pd.to_numeric(
        df["% de Maturation"].astype(str).str.replace("%", "", regex=False).str.strip(),
        errors="coerce"
    ).fillna(0)

    # Dates
    for col in ["Date de dernière activité", "Date de prochaine activité"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    # Pipeline pondéré
    df["Pipeline_pondere"] = df["Revenu_M"] * df["Matu_num"] / 100

    return df

df_raw = load_data()

# Colonnes à masquer dans le tableau
HIDDEN_COLS = ["Commentaire", "Commentaire2", "Commentaire3", "Commentaire4",
               "Activités", "compteur", "Ticket", "Revenu attendu"]

# ─── SIDEBAR FILTRES ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="sidebar-logo">🔷 Exergon</div>', unsafe_allow_html=True)
    st.markdown("### Filtres")

    all_typologies = sorted(df_raw["Typologie"].dropna().unique())
    sel_typo = st.multiselect("Typologie", all_typologies, default=all_typologies)

    all_etapes = sorted(df_raw["Étape"].dropna().unique())
    sel_etape = st.multiselect("Étape", all_etapes, default=all_etapes)

    all_interet = sorted(df_raw["Intérêt"].dropna().unique())
    sel_interet = st.multiselect("Intérêt", all_interet, default=all_interet)

    rev_min, rev_max = float(df_raw["Revenu_M"].min()), float(df_raw["Revenu_M"].max())
    if rev_max > 0:
        rev_range = st.slider("Revenu attendu (M€)", rev_min, rev_max, (rev_min, rev_max), step=0.5)
    else:
        rev_range = (rev_min, rev_max)

    tri_options = sorted(df_raw["Tri"].dropna().unique())
    sel_tri = st.multiselect("Tri activités", tri_options, default=tri_options)

    st.markdown("---")

# ─── FILTRAGE ───────────────────────────────────────────────────────────────────
df = df_raw.copy()
if sel_typo:   df = df[df["Typologie"].isin(sel_typo)]
if sel_etape:  df = df[df["Étape"].isin(sel_etape)]
if sel_interet: df = df[df["Intérêt"].isin(sel_interet)]
if sel_tri:    df = df[df["Tri"].isin(sel_tri)]
df = df[df["Revenu_M"].between(rev_range[0], rev_range[1])]

# ─── HEADER ──────────────────────────────────────────────────────────────────────
today = datetime.today()
st.markdown(f"""
<div class="main-header">
    <h1>📊 Tableau de Bord Pipeline — Exergon</h1>
    <p>Mis à jour le {today.strftime("%d %B %Y")} &nbsp;·&nbsp; {len(df)} leads filtrés sur {len(df_raw)} au total</p>
</div>
""", unsafe_allow_html=True)

# ─── ONGLETS ────────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📈  KPIs & Analyse", "📋  Pipeline Détaillé"])

# ════════════════════════════════════════════════════════════════════════════════
# TAB 1 — KPIs
# ════════════════════════════════════════════════════════════════════════════════
with tab1:

    # KPI helpers
    total_rev    = df["Revenu_M"].sum()
    total_leads  = len(df)
    avg_ticket   = df[df["Revenu_M"] > 0]["Revenu_M"].mean() if (df["Revenu_M"] > 0).any() else 0
    pondere      = df["Pipeline_pondere"].sum()
    matu_moy     = df["Matu_num"].mean()
    ko_count     = (df["Intérêt"] == "KO").sum()
    ok_count     = (df["Intérêt"] == "OK").sum()
    conv_rate    = ok_count / total_leads * 100 if total_leads else 0
    mort_count   = df["Étape"].isin(["Mort / Plus rien à faire","Probablement mort"]).sum()
    actif_count  = total_leads - mort_count

    now = pd.Timestamp.now()
    en_retard = df[
        df["Date de prochaine activité"].notna() &
        (df["Date de prochaine activité"] < now)
    ].shape[0]

    # ── Ligne 1 KPIs ────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">Vue d\'ensemble</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">💰 Pipeline Brut</div>
            <div class="kpi-value">{total_rev:.1f} M€</div>
            <div class="kpi-sub">{total_leads} leads au total</div>
        </div>""", unsafe_allow_html=True)

    with c2:
        badge = "badge-green" if pondere >= 5 else "badge-orange" if pondere >= 1 else "badge-red"
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">🎯 Pipeline Pondéré</div>
            <div class="kpi-value">{pondere:.2f} M€</div>
            <span class="kpi-badge {badge}">Matu. moy. {matu_moy:.0f}%</span>
        </div>""", unsafe_allow_html=True)

    with c3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">🎟️ Ticket Moyen</div>
            <div class="kpi-value">{avg_ticket:.2f} M€</div>
            <div class="kpi-sub">leads avec revenu &gt; 0</div>
        </div>""", unsafe_allow_html=True)

    with c4:
        badge_conv = "badge-green" if conv_rate >= 30 else "badge-orange" if conv_rate >= 10 else "badge-red"
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">✅ Taux de Conversion</div>
            <div class="kpi-value">{conv_rate:.1f}%</div>
            <span class="kpi-badge {badge_conv}">{ok_count} OK · {ko_count} KO</span>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Ligne 2 KPIs ────────────────────────────────────────────────────────
    c5, c6, c7, c8 = st.columns(4)

    with c5:
        badge_act = "badge-green" if actif_count / total_leads > 0.7 else "badge-orange"
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">🟢 Leads Actifs</div>
            <div class="kpi-value">{actif_count}</div>
            <span class="kpi-badge {badge_act}">{mort_count} morts/perdus</span>
        </div>""", unsafe_allow_html=True)

    with c6:
        badge_ret = "badge-red" if en_retard > 10 else "badge-orange" if en_retard > 5 else "badge-green"
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">⏰ Activités en Retard</div>
            <div class="kpi-value">{en_retard}</div>
            <span class="kpi-badge {badge_ret}">relances à planifier</span>
        </div>""", unsafe_allow_html=True)

    with c7:
        # Leads avec activité cette semaine
        semaine = df["Tri"].isin(["Semaine passée","Semaine passée & à venir"]).sum()
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">📅 Actifs cette Semaine</div>
            <div class="kpi-value">{semaine}</div>
            <div class="kpi-sub">activité semaine passée</div>
        </div>""", unsafe_allow_html=True)

    with c8:
        top_typo = df.groupby("Typologie")["Revenu_M"].sum().idxmax() if total_leads else "—"
        top_val  = df.groupby("Typologie")["Revenu_M"].sum().max() if total_leads else 0
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">🏆 Top Typologie</div>
            <div class="kpi-value" style="font-size:1.1rem;padding-top:0.3rem">{top_typo}</div>
            <span class="kpi-badge badge-blue">{top_val:.1f} M€</span>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Graphiques ──────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">Analyse Graphique</div>', unsafe_allow_html=True)

    col_l, col_r = st.columns(2)

    with col_l:
        # Pipeline par étape
        etape_data = (
            df.groupby("Étape")["Revenu_M"]
            .sum()
            .reset_index()
            .sort_values("Revenu_M", ascending=True)
        )
        fig_etape = px.bar(
            etape_data, x="Revenu_M", y="Étape", orientation="h",
            title="Pipeline par Étape (M€)",
            color="Revenu_M",
            color_continuous_scale=["#c7d2fe","#6366f1","#1a1a2e"],
            labels={"Revenu_M": "Revenu (M€)", "Étape": ""}
        )
        fig_etape.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            coloraxis_showscale=False,
            margin=dict(l=10, r=10, t=40, b=10),
            font_family="Inter",
            title_font_size=14, title_font_color="#1a1a2e"
        )
        fig_etape.update_traces(marker_line_width=0)
        st.plotly_chart(fig_etape, use_container_width=True)

    with col_r:
        # Répartition par Typologie (donut)
        typo_data = df.groupby("Typologie").size().reset_index(name="Nb")
        fig_typo = px.pie(
            typo_data, names="Typologie", values="Nb",
            title="Répartition par Typologie",
            hole=0.5,
            color_discrete_sequence=px.colors.sequential.Blues_r
        )
        fig_typo.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(l=10, r=10, t=40, b=10),
            font_family="Inter",
            title_font_size=14, title_font_color="#1a1a2e",
            legend=dict(font_size=10)
        )
        st.plotly_chart(fig_typo, use_container_width=True)

    col_l2, col_r2 = st.columns(2)

    with col_l2:
        # Maturation distribution
        fig_matu = px.histogram(
            df[df["Matu_num"] > 0], x="Matu_num", nbins=10,
            title="Distribution du Taux de Maturation (%)",
            color_discrete_sequence=["#6366f1"],
            labels={"Matu_num": "Maturation (%)", "count": "Nb leads"}
        )
        fig_matu.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(l=10, r=10, t=40, b=10),
            font_family="Inter",
            title_font_size=14, title_font_color="#1a1a2e",
            bargap=0.1
        )
        st.plotly_chart(fig_matu, use_container_width=True)

    with col_r2:
        # Pipeline pondéré par Typologie
        pond_typo = (
            df.groupby("Typologie")["Pipeline_pondere"]
            .sum()
            .reset_index()
            .sort_values("Pipeline_pondere", ascending=False)
            .head(8)
        )
        fig_pond = px.bar(
            pond_typo, x="Typologie", y="Pipeline_pondere",
            title="Pipeline Pondéré par Typologie (M€)",
            color="Pipeline_pondere",
            color_continuous_scale=["#c7d2fe","#6366f1","#1a1a2e"],
            labels={"Pipeline_pondere": "Pipeline Pondéré (M€)", "Typologie": ""}
        )
        fig_pond.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            coloraxis_showscale=False,
            margin=dict(l=10, r=10, t=40, b=10),
            font_family="Inter",
            title_font_size=14, title_font_color="#1a1a2e",
            xaxis_tickangle=-30
        )
        fig_pond.update_traces(marker_line_width=0)
        st.plotly_chart(fig_pond, use_container_width=True)

    # ── Funnel ──────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">Funnel de Conversion</div>', unsafe_allow_html=True)

    funnel_order = [
        "Intérêt à qualifier",
        "Interêt indicatif - A creuser",
        "Pas encore prêt - A alimenter",
        "Probablement mort",
        "Mort / Plus rien à faire",
    ]
    funnel_counts = [df[df["Étape"] == e].shape[0] for e in funnel_order]
    funnel_vals   = [df[df["Étape"] == e]["Revenu_M"].sum() for e in funnel_order]

    fig_funnel = go.Figure(go.Funnel(
        y=funnel_order,
        x=funnel_counts,
        textinfo="value+percent initial",
        marker=dict(color=["#1a1a2e","#0f3460","#16213e","#d97706","#dc2626"]),
        textfont=dict(color="white", size=13, family="Inter")
    ))
    fig_funnel.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(l=10, r=10, t=20, b=10),
        font_family="Inter", height=320
    )
    st.plotly_chart(fig_funnel, use_container_width=True)


# ════════════════════════════════════════════════════════════════════════════════
# TAB 2 — TABLEAU
# ════════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">Pipeline Détaillé</div>', unsafe_allow_html=True)

    # Filtres inline
    fc1, fc2, fc3 = st.columns([2, 2, 3])
    with fc1:
        search = st.text_input("🔍 Rechercher (entité, contact…)", "")
    with fc2:
        sort_col = st.selectbox("Trier par", ["Entité","Revenu_M","Matu_num","Étape","Typologie"])
    with fc3:
        sort_dir = st.radio("Ordre", ["Décroissant","Croissant"], horizontal=True)

    # Préparation du tableau
    display_cols = [c for c in df.columns if c not in HIDDEN_COLS]
    df_display = df[display_cols].copy()

    # Renommage pour plus de clarté
    df_display = df_display.rename(columns={
        "Revenu_M": "Revenu (M€)",
        "Matu_num": "Maturation (%)",
        "Pipeline_pondere": "Pipeline Pondéré (M€)"
    })

    # Recherche
    if search:
        mask = df_display.apply(
            lambda col: col.astype(str).str.contains(search, case=False, na=False)
        ).any(axis=1)
        df_display = df_display[mask]

    # Tri
    sort_col_mapped = {
        "Revenu_M": "Revenu (M€)",
        "Matu_num": "Maturation (%)"
    }.get(sort_col, sort_col)
    if sort_col_mapped in df_display.columns:
        df_display = df_display.sort_values(
            sort_col_mapped, ascending=(sort_dir == "Croissant")
        )

    # Format des colonnes numériques
    for col in ["Revenu (M€)", "Pipeline Pondéré (M€)"]:
        if col in df_display.columns:
            df_display[col] = df_display[col].round(2)
    if "Maturation (%)" in df_display.columns:
        df_display["Maturation (%)"] = df_display["Maturation (%)"].apply(
            lambda x: f"{x:.0f}%" if pd.notna(x) else ""
        )

    st.caption(f"{len(df_display)} leads affichés")
    st.dataframe(df_display, use_container_width=True, height=520)

    # Export CSV
    csv = df_display.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Exporter en CSV", csv,
        file_name=f"pipeline_exergon_{today.strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )
