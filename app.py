import os
import io
import json
import time
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import psycopg2
from psycopg2 import pool
from psycopg2.extras import RealDictCursor
from pypdf import PdfReader
from docx import Document

st.set_page_config(page_title="OmniLead | Akquise-Evaluation", layout="wide")

# ==============================================================================
# 1. DATENBANK-VERBINDUNG & TABELLEN-INITIALISIERUNG
# ==============================================================================
def get_db_url():
    try:
        u = st.secrets["DATABASE_URL"].strip()
    except Exception:
        u = os.environ.get("DATABASE_URL", "").strip()
    return u.replace("postgres://", "postgresql://", 1) if u.startswith("postgres://") else u

@st.cache_resource
def get_db_pool():
    u = get_db_url()
    if not u:
        raise RuntimeError("DATABASE_URL fehlt in den Streamlit Secrets.")
    return pool.ThreadedConnectionPool(minconn=1, maxconn=6, dsn=u, connect_timeout=15)

def get_conn():
    p, err = get_db_pool(), None
    for _ in range(3):
        c = None
        try:
            c = p.getconn()
            with c.cursor() as cur:
                cur.execute("SELECT 1;")
            return c
        except Exception as e:
            err = e
            if c:
                try:
                    p.putconn(c, close=True)
                except Exception:
                    pass
            time.sleep(0.4)
    raise err

def put_conn(c, broken=False):
    if c:
        try:
            get_db_pool().putconn(c, close=broken)
        except Exception:
            try:
                c.close()
            except Exception:
                pass

@st.cache_resource
def init_db():
    c, brk = get_conn(), False
    try:
        with c.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS acq_approaches (
                    id SERIAL PRIMARY KEY,
                    name TEXT NOT NULL UNIQUE,
                    guide_text TEXT,
                    goal_1 TEXT,
                    goal_2 TEXT,
                    goal_3 TEXT,
                    category TEXT,
                    created_at TIMESTAMPTZ DEFAULT NOW()
                );
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS acq_interactions (
                    id SERIAL PRIMARY KEY,
                    approach_id INTEGER REFERENCES acq_approaches(id) ON DELETE CASCADE,
                    contact_type TEXT,
                    outcome TEXT,
                    failure_reasons JSONB,
                    objections JSONB,
                    performance INTEGER,
                    created_at TIMESTAMPTZ DEFAULT NOW()
                );
            """)
        c.commit()
    except Exception:
        brk = True
        raise
    finally:
        put_conn(c, brk)
    return True

try:
    init_db()
except Exception as e:
    st.error(f"Datenbank-Fehler: {e}")
    st.stop()

# ==============================================================================
# 2. DATENBANK-OPERATIONEN
# ==============================================================================
def lade_ansaetze():
    c, brk = get_conn(), False
    try:
        with c.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute("SELECT * FROM acq_approaches ORDER BY created_at ASC;")
            return [dict(r) for r in cur.fetchall()]
    except Exception:
        brk = True
        raise
    finally:
        put_conn(c, brk)

def lade_interaktionen(approach_id=None):
    c, brk = get_conn(), False
    try:
        with c.cursor(cursor_factory=RealDictCursor) as cur:
            if approach_id:
                cur.execute("SELECT * FROM acq_interactions WHERE approach_id = %s ORDER BY created_at ASC;", (approach_id,))
            else:
                cur.execute("SELECT * FROM acq_interactions ORDER BY created_at ASC;")
            return [dict(r) for r in cur.fetchall()]
    except Exception:
        brk = True
        raise
    finally:
        put_conn(c, brk)

def speichere_ansatz(name, guide, g1, g2, g3, category):
    c, brk = get_conn(), False
    try:
        with c.cursor() as cur:
            cur.execute("""
                INSERT INTO acq_approaches (name, guide_text, goal_1, goal_2, goal_3, category)
                VALUES (%s, %s, %s, %s, %s, %s) ON CONFLICT (name) DO NOTHING RETURNING id;
            """, (name.strip(), guide, g1, g2, g3, category))
            res = cur.fetchone()
        c.commit()
        return res is not None
    except Exception:
        brk = True
        raise
    finally:
        put_conn(c, brk)

def speichere_interaktion(app_id, contact, outcome, reasons, objections, perf):
    c, brk = get_conn(), False
    try:
        with c.cursor() as cur:
            cur.execute("""
                INSERT INTO acq_interactions (approach_id, contact_type, outcome, failure_reasons, objections, performance)
                VALUES (%s, %s, %s, %s, %s, %s);
            """, (app_id, contact, outcome, json.dumps(reasons), json.dumps(objections), perf))
        c.commit()
    except Exception:
        brk = True
        raise
    finally:
        put_conn(c, brk)

def loesche_ansatz(app_id):
    c, brk = get_conn(), False
    try:
        with c.cursor() as cur:
            cur.execute("DELETE FROM acq_approaches WHERE id = %s;", (app_id,))
        c.commit()
    except Exception:
        brk = True
        raise
    finally:
        put_conn(c, brk)

def extrahiere_leitfaden(uploaded_file):
    if not uploaded_file:
        return ""
    name = uploaded_file.name.lower()
    raw = uploaded_file.read()
    if name.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(raw))
        return "\n".join(p.extract_text() or "" for p in reader.pages).strip()
    if name.endswith(".docx"):
        doc = Document(io.BytesIO(raw))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for tb in doc.tables:
            for r in tb.rows:
                parts.append(" | ".join(c.text.strip() for c in r.cells))
        return "\n".join(parts).strip()
    for enc in ("utf-8", "latin-1"):
        try:
            return raw.decode(enc).strip()
        except Exception:
            pass
    return ""

# ==============================================================================
# 3. DESIGN & CSS (OmniLead Dark & Plastisch)
# ==============================================================================
st.markdown("""
<style>
    .stApp {
        background: radial-gradient(circle at 50% -10%, #0F1A2E 0%, #07090E 55%, #05070B 100%);
        color: #F1F5F9;
    }
    .block-container {
        max-width: 1240px !important;
        padding-top: 3.6rem !important;
        padding-bottom: 3.4rem !important;
        padding-left: 2.4rem !important;
        padding-right: 2.4rem !important;
    }
    hr {
        border: none !important;
        border-top: 1px solid rgba(125,211,252,.22) !important;
        margin: 1.35rem 0 !important;
    }

    /* Tabs Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.75rem;
        background: #0A0F1A;
        padding: 0.45rem;
        border-radius: 14px;
        border: 1px solid rgba(125,211,252,.22);
        box-shadow: inset 0 2px 6px rgba(0,0,0,.55);
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 10px !important;
        padding: 0.62rem 1.35rem !important;
        font-weight: 600 !important;
        font-size: 0.96rem !important;
        color: #94A3B8 !important;
        border: none !important;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(160deg, #162238, #0F172A) !important;
        color: #E0F2FE !important;
        box-shadow: 0 4px 12px rgba(0,0,0,.4), inset 0 1px 0 rgba(125,211,252,.35) !important;
    }

    /* Header */
    .omni-header {
        padding-bottom: 0.85rem;
        margin-bottom: 1.15rem;
        border-bottom: 1.5px solid rgba(125, 211, 252, 0.4);
    }
    .omni-brand {
        font-size: 2.3rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        background: linear-gradient(135deg, #FFFFFF 0%, #E0F2FE 55%, #7DD3FC 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        line-height: 1.15;
    }
    .omni-sub {
        color: #94A3B8;
        font-size: 0.88rem;
        margin-top: 0.22rem;
    }

    /* Sculpt Panel (Masken) */
    .sculpt-panel {
        background: linear-gradient(165deg, #11192B, #0A0F1A);
        border: 1px solid rgba(125,211,252,.26);
        border-radius: 16px;
        padding: 1.65rem 1.8rem;
        box-shadow: 0 14px 30px rgba(0,0,0,.52), inset 0 1px 0 rgba(255,255,255,.09);
    }

    /* Große wortlose Kennzahlen rechts neben der Maske */
    .huge-metric-box {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        min-height: 470px;
        gap: 3.1rem;
        background: linear-gradient(165deg, #0E1525, #080C15);
        border: 1px solid rgba(125,211,252,.2);
        border-radius: 18px;
        padding: 2rem 1.2rem;
        box-shadow: 0 14px 30px rgba(0,0,0,.48), inset 0 1px 0 rgba(255,255,255,.06);
    }
    .huge-val {
        font-size: 5.4rem;
        font-weight: 900;
        line-height: 1;
        letter-spacing: -0.03em;
    }
    .val-blue {
        background: linear-gradient(135deg, #BAE6FD 0%, #38BDF8 60%, #0284C7 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .val-green {
        background: linear-gradient(135deg, #A7F3D0 0%, #34D399 60%, #059669 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .val-stars {
        font-size: 3.1rem;
        letter-spacing: 0.14em;
        color: #FFFFFF;
        text-shadow: 0 0 18px rgba(255,255,255,0.45);
        line-height: 1;
    }

    /* Kategorie-Karten (Eis, Wald, Feuer, Luft) */
    .cat-header {
        font-size: 1.15rem;
        font-weight: 800;
        letter-spacing: 0.03em;
        padding: 0.45rem 0.85rem;
        border-radius: 10px;
        margin-bottom: 0.85rem;
        text-align: center;
    }
    .hdr-Eis { background: rgba(56, 189, 248, 0.16); border: 1px solid #38BDF8; color: #7DD3FC; }
    .hdr-Wald { background: rgba(74, 222, 128, 0.16); border: 1px solid #4ADE80; color: #86EFAC; }
    .hdr-Feuer { background: rgba(248, 113, 113, 0.16); border: 1px solid #F87171; color: #FCA5A5; }
    .hdr-Luft { background: rgba(248, 250, 252, 0.14); border: 1px solid #F8FAFC; color: #FFFFFF; }

    .cat-card {
        border-radius: 14px;
        padding: 1.1rem 1.2rem;
        margin-bottom: 0.45rem;
        box-shadow: 0 8px 20px rgba(0,0,0,0.42), inset 0 1px 0 rgba(255,255,255,0.14);
        border: 1.5px solid rgba(255,255,255,0.15);
    }
    .cat-Eis { background: linear-gradient(145deg, #0C4A6E, #075985); border-color: #38BDF8; color: #F0F9FF; }
    .cat-Wald { background: linear-gradient(145deg, #14532D, #166534); border-color: #4ADE80; color: #F0FDF4; }
    .cat-Feuer { background: linear-gradient(145deg, #7F1D1D, #991B1B); border-color: #F87171; color: #FEF2F2; }
    .cat-Luft { background: linear-gradient(145deg, #E2E8F0, #CBD5E1); border-color: #FFFFFF; color: #0F172A; }

    .cat-title { font-size: 1.08rem; font-weight: 800; margin-bottom: 0.25rem; }
    .cat-rate { font-size: 1.75rem; font-weight: 900; line-height: 1.1; }

    /* Horizontale Radio-Buttons */
    div[data-testid="stRadio"] > div[role="radiogroup"] { gap: 0.7rem; }
    div[data-testid="stRadio"] > div[role="radiogroup"] > label {
        background: linear-gradient(165deg,#101829,#0A0F1A) !important;
        border: 1px solid rgba(125,211,252,.3) !important;
        border-radius: 10px !important;
        padding: 0.55rem 0.95rem !important;
        box-shadow: 0 4px 10px rgba(0,0,0,.4) !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] > label:hover {
        border-color: #7DD3FC !important;
    }

    /* Buttons */
    .stButton > button {
        background: linear-gradient(180deg, #152036 0%, #0D1424 100%) !important;
        border: 1px solid rgba(125, 211, 252, 0.42) !important;
        color: #E0F2FE !important;
        border-radius: 11px !important;
        font-weight: 600 !important;
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.42), inset 0 1px 0 rgba(255, 255, 255, 0.1) !important;
    }
    .stButton > button:hover {
        border-color: #7DD3FC !important;
        background: linear-gradient(180deg, #1B2A47 0%, #111B30 100%) !important;
        color: #FFFFFF !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="omni-header">
    <div class="omni-brand">OmniLead</div>
    <div class="omni-sub">Erfolgsevaluation & Akquise-Analyse</div>
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# 4. FESTE AUSWAHLLISTEN & KATEGORIEN
# ==============================================================================
REASONS_LIST = [
    "Begrüßung die Fragen offen ließ",
    "Vertriebliche oder unsichere Betonung",
    "Reaktanz des Kunden",
    "Fehlende Worte",
    "Unzureichende Nutzenargumentation",
    "Irrelevanter Nutzen",
    "Empfangsdame standhaft",
    "Vorwand nicht behandelt",
    "Einwand nicht behandelt",
    "Kunden genervt",
    "Nicht ausreichend zugehört",
]

OBJECTIONS_LIST = [
    "Kein Bedarf / Interesse",
    "Keine Zeit",
    "Kein Geld / Zu teuer",
    "Angebot schicken",
    "Termin erst in Wochen möglich",
    "Ich verstehe das Angebot nicht",
    "Sie wollen nur Verkaufen",
    "Das Können wir selbst",
    "Wir haben einen Partner",
    "Weiterbildungen sind Ressourcenverschwendung",
    "Mitarbeiter wollen nicht",
    "Kann ich alleine nicht entscheiden",
    "Ich darf dich nicht durchstellen",
]

CATEGORIES = {
    "Eis": "Blau",
    "Wald": "Grün",
    "Feuer": "Rot",
    "Luft": "Weiß",
}

if "last_ansatz" not in st.session_state:
    st.session_state.last_ansatz = None
if "view_approach_id" not in st.session_state:
    st.session_state.view_approach_id = None

alle_ansaetze = lade_ansaetze()
alle_interaktionen = lade_interaktionen()

tab1, tab2, tab3, tab4 = st.tabs(["Erhebung", "Ansatz erstellen", "Ansatzübersicht", "Gesamtstatistik"])

# ==============================================================================
# TAB 1: ERHEBUNG
# ==============================================================================
with tab1:
    if not alle_ansaetze:
        st.info("Bitte erstelle zuerst einen Ansatz im Reiter **Ansatz erstellen**.")
    else:
        col_form, col_metrics = st.columns([2.15, 1.15], gap="large")

        with col_form:
            st.markdown('<div class="sculpt-panel">', unsafe_allow_html=True)

            ansatz_namen = [a["name"] for a in alle_ansaetze]
            def_idx = ansatz_namen.index(st.session_state.last_ansatz) if st.session_state.last_ansatz in ansatz_namen else 0

            sel_ansatz = st.selectbox("Ansatz", ansatz_namen, index=def_idx)
            st.session_state.last_ansatz = sel_ansatz
            sel_app_obj = next(a for a in alle_ansaetze if a["name"] == sel_ansatz)
            sel_app_id = sel_app_obj["id"]

            contact = st.radio("Ansprechpartner", ["Geschäftsführer", "Führungskraft", "Empfang"], horizontal=True)
            outcome = st.radio("Ergebnis", ["Erfolg", "Plan B", "Misserfolg"], horizontal=True)

            reasons = []
            if outcome == "Misserfolg":
                with st.expander("Gründe für Misserfolg auswählen", expanded=True):
                    reasons = st.multiselect("Warum ist es gescheitert?", REASONS_LIST)

            with st.expander("Vor- und Einwände"):
                objections = st.multiselect("Aufgetretene Vor- und Einwände:", OBJECTIONS_LIST)
                custom_obj = st.text_input("Weitere Vor- oder Einwände hinzufügen (kommagetrennt):", placeholder="z. B. Interner Projektstopp, Urlaubszeit...")

            perf = st.radio("Eigene Performance (1 – 5)", [1, 2, 3, 4, 5], index=2, horizontal=True)

            st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
            if st.button("💾 Speichern", use_container_width=True):
                final_obj = list(objections)
                if custom_obj.strip():
                    final_obj.extend([o.strip() for o in custom_obj.split(",") if o.strip()])
                speichere_interaktion(sel_app_id, contact, outcome, reasons, final_obj, perf)
                st.success("Interaktion gespeichert!")
                time.sleep(0.5)
                st.rerun()

            st.markdown('</div>', unsafe_allow_html=True)

        with col_metrics:
            ansatz_interactions = [i for i in alle_interaktionen if i["approach_id"] == sel_app_id]
            total_attempts = len(ansatz_interactions)
            success_count = sum(1 for i in ansatz_interactions if i["outcome"] == "Erfolg")
            sr_pct = round((success_count / total_attempts * 100)) if total_attempts > 0 else 0

            avg_perf = (sum(int(i["performance"]) for i in ansatz_interactions) / total_attempts) if total_attempts > 0 else 0
            star_count = round(avg_perf)
            stars_html = ("★" * star_count + "☆" * (5 - star_count)) if total_attempts > 0 else "☆☆☆☆☆"

            st.markdown(f"""
            <div class="huge-metric-box">
                <div class="huge-val val-blue">{total_attempts}</div>
                <div class="huge-val val-green">{sr_pct}%</div>
                <div class="val-stars">{stars_html}</div>
            </div>
            """, unsafe_allow_html=True)

# ==============================================================================
# TAB 2: ANSATZ ERSTELLEN
# ==============================================================================
with tab2:
    st.markdown('<div class="sculpt-panel">', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    new_name = c1.text_input("Ansatz benennen", placeholder="z. B. Direktansprache Produktivität...")
    new_cat = c2.selectbox(
        "Farbe / Kategorie zuweisen",
        ["Eis", "Wald", "Feuer", "Luft"],
        format_func=lambda x: f"{CATEGORIES[x]} ({x})",
    )

    new_guide = st.file_uploader("Leitfaden hochladen (PDF, DOCX, TXT)", type=["pdf", "docx", "txt"])
    preview_text = extrahiere_leitfaden(new_guide) if new_guide else ""
    if preview_text:
        with st.expander("📄 Hochgeladenen Leitfaden einsehen", expanded=True):
            st.text_area("Inhalt des Leitfadens", value=preview_text, height=200, disabled=True)

    g1 = st.text_input("Ziel", placeholder="Hauptziel dieses Ansatzes...")
    g2 = st.text_input("Zweitziel", placeholder="Sekundäres Ziel...")
    g3 = st.text_input("Ausweichziel", placeholder="Plan B / Fallback-Ziel...")

    if st.button("💾 Ansatz speichern", use_container_width=True):
        if not new_name.strip():
            st.warning("Bitte gib dem Ansatz einen Namen.")
        else:
            if speichere_ansatz(new_name, preview_text, g1, g2, g3, new_cat):
                st.success(f"Ansatz **{new_name.strip()}** wurde angelegt!")
                time.sleep(0.5)
                st.rerun()
            else:
                st.warning("Ein Ansatz mit diesem Namen existiert bereits.")
    st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# STATISTIK-AUFBEREITUNG (KURS- & BALKEN-DIAGRAMME)
# ==============================================================================
def render_stats(interactions_list, title_prefix=""):
    if not interactions_list:
        st.info("Für diese Auswahl wurden noch keine Interaktionen erfasst.")
        return

    df = pd.DataFrame(interactions_list).sort_values("created_at")
    df["Anwendungsfall"] = range(1, len(df) + 1)
    df["Kum_Erfolge"] = df["outcome"].apply(lambda x: 1 if x == "Erfolg" else 0).cumsum()
    df["Kum_Misserfolge"] = df["outcome"].apply(lambda x: 1 if x == "Misserfolg" else 0).cumsum()

    fig_line = go.Figure()
    fig_line.add_trace(go.Scatter(
        x=df["Anwendungsfall"],
        y=df["Kum_Erfolge"],
        mode="lines+markers",
        name="Erfolge",
        line=dict(color="#38BDF8", width=3.5),
        marker=dict(size=8),
    ))
    fig_line.add_trace(go.Scatter(
        x=df["Anwendungsfall"],
        y=df["Kum_Misserfolge"],
        mode="lines+markers",
        name="Misserfolge",
        line=dict(color="#F87171", width=3.5),
        marker=dict(size=8),
    ))
    fig_line.update_layout(
        title=f"{title_prefix}Verlaufskurve (Erfolge & Misserfolge)",
        template="plotly_dark",
        plot_bgcolor="rgba(10,15,26,0.65)",
        paper_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Anwendungsfälle (Anzahl)",
        yaxis_title="Erfolge / Misserfolge",
        xaxis=dict(dtick=1),
        yaxis=dict(dtick=1),
        hovermode="x unified",
    )
    st.plotly_chart(fig_line, use_container_width=True)

    def plot_grouped_bars(col_name, chart_title):
        records = []
        for _, row in df.iterrows():
            items = row[col_name]
            if isinstance(items, str):
                try:
                    items = json.loads(items)
                except Exception:
                    items = []
            if not isinstance(items, list):
                items = []
            for item in items:
                records.append({"Ansprechpartner": row["contact_type"], "Kriterium": item})

        if not records:
            st.caption(f"Noch keine Einträge für: {chart_title}")
            return

        df_exp = pd.DataFrame(records)
        df_grp = df_exp.groupby(["Kriterium", "Ansprechpartner"]).size().reset_index(name="Häufigkeit")

        fig = px.bar(
            df_grp,
            x="Kriterium",
            y="Häufigkeit",
            color="Ansprechpartner",
            barmode="group",
            text="Häufigkeit",
            color_discrete_map={
                "Geschäftsführer": "#38BDF8",
                "Führungskraft": "#4ADE80",
                "Empfang": "#FACC15",
            },
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(
            title=chart_title,
            template="plotly_dark",
            plot_bgcolor="rgba(10,15,26,0.65)",
            paper_bgcolor="rgba(0,0,0,0)",
            xaxis_title="",
            yaxis_title="Häufigkeit",
            yaxis=dict(dtick=1),
        )
        st.plotly_chart(fig, use_container_width=True)

    c_left, c_right = st.columns(2)
    with c_left:
        plot_grouped_bars("failure_reasons", "Gründe für Misserfolg (nach Ansprechpartner)")
    with c_right:
        plot_grouped_bars("objections", "Vor- und Einwände (nach Ansprechpartner)")

# ==============================================================================
# TAB 3: ANSATZÜBERSICHT
# ==============================================================================
with tab3:
    if not alle_ansaetze:
        st.info("Noch keine Ansätze vorhanden.")
    else:
        cols = st.columns(4, gap="medium")
        cat_order = ["Eis", "Wald", "Feuer", "Luft"]

        for idx, cat in enumerate(cat_order):
            with cols[idx]:
                st.markdown(f'<div class="cat-header hdr-{cat}">{cat}</div>', unsafe_allow_html=True)
                cat_approaches = [a for a in alle_ansaetze if a["category"] == cat]
                if not cat_approaches:
                    st.caption("Keine Ansätze")
                for a in cat_approaches:
                    a_inters = [i for i in alle_interaktionen if i["approach_id"] == a["id"]]
                    t_cnt = len(a_inters)
                    sr = round((sum(1 for i in a_inters if i["outcome"] == "Erfolg") / t_cnt * 100)) if t_cnt > 0 else 0

                    st.markdown(f"""
                    <div class="cat-card cat-{cat}">
                        <div class="cat-title">{a['name']}</div>
                        <div class="cat-rate">{sr}%</div>
                    </div>
                    """, unsafe_allow_html=True)
                    if st.button("Statistik öffnen", key=f"btn_anl_{a['id']}", use_container_width=True):
                        st.session_state.view_approach_id = a["id"]
                    st.markdown("<div style='height:0.45rem'></div>", unsafe_allow_html=True)

        if st.session_state.view_approach_id:
            sel_a = next((a for a in alle_ansaetze if a["id"] == st.session_state.view_approach_id), None)
            if sel_a:
                st.markdown("---")
                st.markdown(f"### 📊 Statistik & Details: {sel_a['name']} ({sel_a['category']})")

                with st.expander("📄 Leitfaden & definierte Ziele einsehen", expanded=False):
                    g_col1, g_col2, g_col3 = st.columns(3)
                    g_col1.markdown(f"**Ziel:** {sel_a.get('goal_1') or '–'}")
                    g_col2.markdown(f"**Zweitziel:** {sel_a.get('goal_2') or '–'}")
                    g_col3.markdown(f"**Ausweichziel:** {sel_a.get('goal_3') or '–'}")
                    if sel_a.get("guide_text"):
                        st.text_area("Hinterlegter Leitfaden", value=sel_a["guide_text"], height=180, disabled=True)
                    else:
                        st.caption("Für diesen Ansatz wurde kein Leitfaden hochgeladen.")

                a_data = [i for i in alle_interaktionen if i["approach_id"] == sel_a["id"]]
                render_stats(a_data, title_prefix=f"{sel_a['name']} – ")

                with st.expander("Gefahrenzone: Ansatz löschen"):
                    if st.button(f"Ansatz '{sel_a['name']}' endgültig löschen", key=f"del_app_{sel_a['id']}"):
                        loesche_ansatz(sel_a["id"])
                        st.session_state.view_approach_id = None
                        st.rerun()

# ==============================================================================
# TAB 4: GESAMTSTATISTIK
# ==============================================================================
with tab4:
    st.markdown("### 🌍 Gesamtstatistik (Alle Ansätze kumuliert)")
    render_stats(alle_interaktionen, title_prefix="Gesamt – ")