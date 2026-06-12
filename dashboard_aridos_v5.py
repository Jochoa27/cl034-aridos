"""
dashboard_aridos_v5.py — V5 (Apache ECharts · diseño showcase)
Puerto: py -m streamlit run dashboard_aridos_v5.py --server.port 8505
Base: V3 — misma lógica de datos, gráficos reemplazados con ECharts 5 via HTML
"""

import streamlit as st
import pandas as pd
from pathlib import Path
import json
from datetime import datetime
from streamlit.components.v1 import html as st_html
from streamlit_autorefresh import st_autorefresh

def _detectar_proyecto() -> tuple[Path, dict]:
    ROOT = Path(__file__).parent / "proyectos"
    codigo = st.query_params.get("proyecto", "")
    carpeta = None
    if codigo and ROOT.exists():
        matches = [d for d in ROOT.iterdir() if d.is_dir() and d.name.startswith(codigo)]
        carpeta = matches[0] if matches else None
    if carpeta is None and ROOT.exists():
        projs = sorted(d for d in ROOT.iterdir() if d.is_dir() and (d / "proyecto.json").exists())
        carpeta = projs[0] if projs else None
    if carpeta and (carpeta / "proyecto.json").exists():
        cfg = json.loads((carpeta / "proyecto.json").read_text(encoding="utf-8"))
        return carpeta, cfg
    return Path(__file__).parent, {"codigo": "", "nombre": "", "archivo_ppto": "Ppto.xlsx"}

_CARPETA_PROYECTO, _CFG = _detectar_proyecto()
PROYECTO_CODIGO = _CFG.get("codigo", "")
PROYECTO_NOMBRE = _CFG.get("nombre", "")

st.set_page_config(
    page_title=f"Control Áridos {PROYECTO_CODIGO} — Constructora Londres",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS PREMIUM ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
.stApp { background-color: #060B15 !important; }
[data-testid="stHeader"] { background: transparent !important; }
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#0A1122 0%,#0C1628 100%) !important;
    border-right: 1px solid rgba(59,130,246,0.13) !important;
}
h1 { font-size:1.65rem !important; font-weight:900 !important; color:#F8FAFC !important; letter-spacing:0.02em !important; }
h2 { font-size:0.80rem !important; font-weight:700 !important; letter-spacing:0.13em !important; text-transform:uppercase !important; color:#475569 !important; border-left:none !important; padding-left:0 !important; margin-top:1.8rem !important; margin-bottom:0.25rem !important; border:none !important; }
h3 { font-size:0.78rem !important; font-weight:700 !important; letter-spacing:0.09em !important; text-transform:uppercase !important; color:#334155 !important; margin-top:0.8rem !important; }
h5 { font-size:0.78rem !important; font-weight:700 !important; letter-spacing:0.08em !important; text-transform:uppercase !important; color:#64748B !important; }
[data-testid="metric-container"] { background:rgba(255,255,255,0.030) !important; border:1px solid rgba(255,255,255,0.070) !important; border-radius:14px !important; padding:16px 18px !important; transition: border-color .2s, box-shadow .2s !important; }
[data-testid="metric-container"]:hover { border-color:rgba(56,189,248,0.30) !important; box-shadow:0 4px 20px rgba(56,189,248,0.10) !important; }
[data-testid="stMetricValue"]  { font-size:1.55rem !important; font-weight:800 !important; color:#F1F5F9 !important; }
[data-testid="stMetricLabel"]  { font-size:0.65rem !important; font-weight:700 !important; letter-spacing:0.10em !important; text-transform:uppercase !important; color:#475569 !important; }
[data-testid="stMetricDelta"]  { font-size:0.74rem !important; font-weight:600 !important; }
[data-testid="stTabs"] > div:first-child { border-bottom:1px solid rgba(255,255,255,0.07) !important; }
button[data-baseweb="tab"] { font-size:0.82rem !important; font-weight:600 !important; letter-spacing:0.04em !important; color:#64748B !important; padding:8px 16px !important; }
button[data-baseweb="tab"][aria-selected="true"] { color:#38BDF8 !important; font-weight:700 !important; }
hr { border:none !important; border-top:1px solid rgba(255,255,255,0.065) !important; margin:1.2rem 0 !important; }
[data-testid="stDataFrame"] { border:1px solid rgba(255,255,255,0.07) !important; border-radius:12px !important; overflow:hidden !important; }
[data-testid="stPopover"] > button { background:rgba(56,189,248,0.04) !important; border:1px solid rgba(56,189,248,0.15) !important; border-radius:8px !important; color:#94A3B8 !important; font-size:0.83rem !important; font-weight:500 !important; }
[data-testid="stPopover"] > button:hover { border-color:rgba(56,189,248,0.35) !important; background:rgba(56,189,248,0.09) !important; }
div[data-testid="stCheckbox"] label { font-size:0.82rem !important; padding-top:1px !important; padding-bottom:1px !important; }
div[data-testid="stCheckbox"] { margin-bottom:-4px !important; }
div[data-testid="stTextInput"] input { font-size:0.83rem !important; }
section[data-testid="stMain"] [data-testid="stButton"] > button { height:auto !important; min-height:78px !important; white-space:pre-line !important; text-align:center !important; line-height:1.65 !important; padding:12px 10px !important; font-size:0.80rem !important; font-weight:600 !important; letter-spacing:0.03em !important; background:rgba(255,255,255,0.030) !important; border:1px solid rgba(255,255,255,0.072) !important; border-radius:14px !important; color:#94A3B8 !important; transition:all .2s !important; }
section[data-testid="stMain"] [data-testid="stButton"] > button:hover { border-color:rgba(56,189,248,0.35) !important; background:rgba(56,189,248,0.07) !important; color:#E2E8F0 !important; }
section[data-testid="stMain"] [data-testid="stButton"] > button[data-testid="baseButton-primary"] { background:rgba(56,189,248,0.14) !important; border:1.5px solid #38BDF8 !important; color:#BAE6FD !important; box-shadow:0 0 22px rgba(56,189,248,0.20) !important; }
.stat-wrap { border-radius:18px; padding:22px 24px; text-align:center; }
.stat-critico { background:linear-gradient(135deg,rgba(255,71,87,0.20),rgba(255,71,87,0.04)); border:1.5px solid rgba(255,71,87,0.55); box-shadow:0 0 60px rgba(255,71,87,0.22), inset 0 0 30px rgba(255,71,87,0.06); }
.stat-riesgo  { background:linear-gradient(135deg,rgba(255,179,0,0.20),rgba(255,179,0,0.04));   border:1.5px solid rgba(255,179,0,0.55);  box-shadow:0 0 60px rgba(255,179,0,0.22),  inset 0 0 30px rgba(255,179,0,0.06); }
.stat-normal  { background:linear-gradient(135deg,rgba(35,209,96,0.20),rgba(35,209,96,0.04));   border:1.5px solid rgba(35,209,96,0.55);  box-shadow:0 0 60px rgba(35,209,96,0.22),  inset 0 0 30px rgba(35,209,96,0.06); }
.stat-lbl { font-size:0.60rem; font-weight:700; letter-spacing:0.18em; text-transform:uppercase; color:rgba(255,255,255,0.40); margin-bottom:10px; }
.stat-val { font-size:2.9rem; font-weight:900; line-height:1.0; margin-bottom:8px; letter-spacing:0.04em; }
.stat-sub { font-size:0.75rem; color:rgba(255,255,255,0.58); line-height:1.45; }
.kpi-g { display:grid; gap:10px; }
.kpi-c { background:linear-gradient(135deg,rgba(56,189,248,0.07),rgba(56,189,248,0.02)); border:1px solid rgba(56,189,248,0.16); border-radius:14px; padding:16px 16px 12px; text-align:center; transition:border-color .2s, box-shadow .2s; }
.kpi-c:hover { border-color:rgba(56,189,248,0.38); box-shadow:0 4px 22px rgba(56,189,248,0.12); }
.kpi-lb { font-size:0.60rem; font-weight:700; letter-spacing:0.13em; text-transform:uppercase; color:#475569; margin-bottom:7px; }
.kpi-vl { font-size:1.75rem; font-weight:800; color:#F1F5F9; line-height:1.0; margin-bottom:5px; font-variant-numeric:tabular-nums; }
.kpi-vls { font-size:1.42rem; font-weight:800; color:#F1F5F9; line-height:1.0; margin-bottom:5px; font-variant-numeric:tabular-nums; }
.dp { font-size:0.73rem; font-weight:600; color:#23D160; }
.dn { font-size:0.73rem; font-weight:600; color:#FF4757; }
.d0 { font-size:0.73rem; font-weight:600; color:#475569; }
.row-lbl { font-size:0.68rem; font-weight:800; letter-spacing:0.10em; text-transform:uppercase; color:#64748B; margin-bottom:7px; }
.sh { display:flex; align-items:center; gap:14px; margin:2.2rem 0 0.9rem; padding-bottom:13px; border-bottom:1px solid rgba(56,189,248,0.13); }
.sh-bar { width:5px; height:30px; border-radius:3px; flex-shrink:0; }
.sh-icon { width:36px; height:36px; border-radius:10px; display:flex; align-items:center; justify-content:center; font-size:1.15rem; flex-shrink:0; }
.sh-txt { font-size:1.05rem; font-weight:800; letter-spacing:0.10em; text-transform:uppercase; color:#E2E8F0; text-shadow:0 0 22px rgba(56,189,248,0.30); }
.sh-badge { margin-left:auto; font-size:0.67rem; font-weight:700; letter-spacing:0.06em; padding:4px 14px; border-radius:20px; border:1px solid rgba(56,189,248,0.25); }
.mk { background:linear-gradient(135deg,rgba(56,189,248,0.07),rgba(56,189,248,0.02)); border:1px solid rgba(56,189,248,0.18); border-radius:12px; padding:12px 14px; transition:border-color .2s,box-shadow .2s; }
.mk-lb { font-size:0.57rem; font-weight:700; letter-spacing:0.11em; text-transform:uppercase; color:#64748B; margin-bottom:4px; }
.mk-vl { font-size:1.28rem; font-weight:800; color:#F1F5F9; line-height:1; }
.mk-sb { font-size:0.67rem; color:#475569; margin-top:2px; }
.risk-cap { font-size:0.78rem; color:#64748B; line-height:1.55; background:rgba(255,255,255,0.018); border:1px solid rgba(255,255,255,0.050); border-radius:10px; padding:10px 14px; margin-bottom:10px; }
.proj-card { background:linear-gradient(135deg,rgba(56,189,248,0.07),rgba(56,189,248,0.01)); border:1px solid rgba(56,189,248,0.22); border-radius:16px; padding:24px 22px; text-align:center; margin-bottom:12px; transition:border-color .2s,box-shadow .2s; }
.proj-card:hover { border-color:rgba(56,189,248,0.55); box-shadow:0 0 28px rgba(56,189,248,0.12); }
/* ── PROGRESS RINGS ─────────────────────────────────────────────────────── */
@keyframes ring-draw { from { stroke-dashoffset: 226.2; } }
.ring-wrap { display:flex; flex-direction:column; align-items:center; gap:5px; padding:18px 8px 14px; background:linear-gradient(135deg,rgba(56,189,248,0.06),rgba(56,189,248,0.01)); border:1px solid rgba(56,189,248,0.15); border-radius:18px; transition:border-color .25s,box-shadow .25s; }
.ring-wrap:hover { border-color:rgba(56,189,248,0.40); box-shadow:0 0 32px rgba(56,189,248,0.15); }
.ring-lbl { font-size:0.57rem; font-weight:700; letter-spacing:0.14em; text-transform:uppercase; color:#475569; text-align:center; }
.ring-val { font-size:1.22rem; font-weight:800; color:#F1F5F9; line-height:1; font-variant-numeric:tabular-nums; }
.ring-sub { font-size:0.64rem; color:#475569; text-align:center; line-height:1.35; }
circle.progress { animation: ring-draw 1.5s cubic-bezier(.4,0,.2,1) both; }
/* ── ENHANCED KPIS ──────────────────────────────────────────────────────── */
.mk:hover { border-color:rgba(56,189,248,0.32); box-shadow:0 0 22px rgba(56,189,248,0.12); }
</style>
""", unsafe_allow_html=True)

# ── SELECTOR MULTI-PROYECTO ───────────────────────────────────────────────────
_PROJS_ROOT = Path(__file__).parent / "proyectos"
_TODOS_PROJS = (
    sorted(d for d in _PROJS_ROOT.iterdir()
           if d.is_dir() and (d / "proyecto.json").exists())
    if _PROJS_ROOT.exists() else []
)
if len(_TODOS_PROJS) > 1 and not st.query_params.get("proyecto", ""):
    st.markdown("""
    <div style="text-align:center;padding:40px 0 20px;">
      <div style="font-size:0.65rem;font-weight:700;letter-spacing:0.22em;text-transform:uppercase;color:#5470c6;margin-bottom:8px;">CONSTRUCTORA LONDRES</div>
      <div style="font-size:2.2rem;font-weight:900;color:#F8FAFC;">🏗️ CONTROL DE ÁRIDOS</div>
      <div style="font-size:0.92rem;color:#475569;margin-top:8px;">Selecciona un proyecto</div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()
    _cols = st.columns(min(len(_TODOS_PROJS), 3), gap="large")
    for _i, _d in enumerate(_TODOS_PROJS):
        _c = json.loads((_d / "proyecto.json").read_text(encoding="utf-8"))
        _n_xlsx = sum(1 for _ in _d.glob("*.xlsx"))
        with _cols[_i % 3]:
            st.markdown(f"""<div class="proj-card">
              <div style="font-size:0.60rem;font-weight:700;letter-spacing:0.18em;text-transform:uppercase;color:#38BDF8;margin-bottom:6px;">OBRA</div>
              <div style="font-size:1.55rem;font-weight:900;color:#F8FAFC;margin-bottom:4px;">{_c.get("codigo","")}</div>
              <div style="font-size:0.88rem;color:#94A3B8;margin-bottom:12px;">{_c.get("nombre","")}</div>
              <div style="font-size:0.70rem;color:#334155;">{_n_xlsx} archivos de datos</div>
            </div>""", unsafe_allow_html=True)
            if st.button("Ver dashboard →", key=f"_sel_{_c.get('codigo','')}", use_container_width=True):
                st.query_params["proyecto"] = _c.get("codigo", "")
                st.rerun()
    st.stop()

# ── ECHARTS ──────────────────────────────────────────────────────────────────
_EC_CDN = "https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"

def ec_html(js_option: str, height: int = 420) -> None:
    tpl = (
        f'<!DOCTYPE html><html><head><meta charset="utf-8">'
        f'<style>body{{margin:0;padding:0;background:transparent;overflow:hidden;}}'
        f'#c{{width:100%;height:{height - 14}px;}}</style></head>'
        f'<body><div id="c"></div>'
        f'<script src="{_EC_CDN}"></script>'
        f'<script>'
        f'var c=echarts.init(document.getElementById("c"),null,{{backgroundColor:"transparent",renderer:"canvas"}});'
        f'c.setOption({js_option});'
        f'window.addEventListener("resize",function(){{c.resize();}});'
        f'</script></body></html>'
    )
    st_html(tpl, height=height)

# ── PALETA ECHARTS ────────────────────────────────────────────────────────────
C_PPTO     = "#5470c6"
C_OC       = "#fc8452"
C_RECIBIDO = "#73c0de"
C_CRITICO  = "#FF4757"
C_ALERTA   = "#FFB300"
C_OK       = "#23D160"
PALETA_CC  = ["#5470c6","#fc8452","#73c0de","#91cc75","#fac858",
              "#9a60b4","#ee6666","#38BDF8","#3ba272","#ea7ccc","#2DD4BF"]

# ── CONSTANTES ────────────────────────────────────────────────────────────────
CARPETA        = _CARPETA_PROYECTO
ARCHIVO_PPTO   = CARPETA / _CFG["archivo_ppto"]
ARCHIVO_OC     = CARPETA / "Descarga_IConstruye.xlsx"
ARCHIVO_RECFAC = CARPETA / "Recepcion_Facturación_OC.xlsx"
ARCHIVO_NC     = CARPETA / "ReporteControlNotasCorreccion.xlsx"
INTERVALO      = 60

PALABRAS_ARIDOS = ["arena","grava","bolón","bolon","integral",
                   "base estabilizada","ripio","gravilla","polvo de piedra",
                   "tierra","maicillo"]
MAPA_NOMBRES = {
    "grava chancada 1.1/2''":                     'Grava Chancada 1½"',
    'grava chancada 1.1/2"':                      'Grava Chancada 1½"',
    "grava chancada":                             "Grava Chancada",
    "arena gruesa":                               "Arena Gruesa",
    "arena fina":                                 "Arena Fina",
    "arena de relleno":                           "Arena de Relleno",
    "base estabilizada chancada 1 ½\" cbr >100%": "Base Estabilizada",
    "base estabilizada cbr >20%":                 "Base Estabilizada",
    "base estabilizada":                          "Base Estabilizada",
    "bolón de 6\" a 10\"":                        'Bolón 6"-10"',
    "bolón":                                      "Bolón",
    "bolon":                                      "Bolón",
    "integral bajo 2":                            'Integral Bajo 2"',
    "integral bajo 3":                            'Integral Bajo 3"',
    "integral bajo 4":                            'Integral Bajo 4"',
    "integral":                                   "Integral",
    "cama de ripio":                              "Ripio (Cama)",
    "maicillo":                                   "Maicillo",
    "tierra / material":                          "Tierra / Mat. Excavación",
    "tierra para relleno":                        "Tierra para Relleno",
    "tierra":                                     "Tierra",
}

# ── HELPERS ───────────────────────────────────────────────────────────────────
def _limpiar(s):
    return str(s).strip().lower().replace('\xa0', ' ').replace(' ', ' ')

def normalizar(nombre):
    n = _limpiar(nombre)
    for k, v in MAPA_NOMBRES.items():
        if k in n: return v
    return str(nombre).strip().title()

def es_arido(texto):
    return any(p in _limpiar(texto) for p in PALABRAS_ARIDOS)

def semaforo(desv, umbral):
    if pd.isna(desv): return "⚪"
    if desv < -umbral:     return "🔴"
    if desv < -umbral / 2: return "🟡"
    return "🟢"

def _n(v, d=0, pfx="", sfx="", sgn=False):
    if pd.isna(v): return "-"
    fmt = f"+,.{d}f" if sgn else f",.{d}f"
    s = f"{v:{fmt}}".replace(".", "\x00").replace(",", ".").replace("\x00", ",")
    return f"{pfx}{s}{sfx}"

_F0  = lambda x: _n(x, 0)
_F1  = lambda x: _n(x, 1)
_F2  = lambda x: _n(x, 2)
_FS0 = lambda x: _n(x, 0, "$")
_FP  = lambda x: "-" if pd.isna(x) else f"{x:.1f}%".replace(".", ",")
_FPp = lambda x: "-" if pd.isna(x) else f"{x:+.1f}%".replace(".", ",")
_FI  = lambda x: "-" if pd.isna(x) or x <= 0 else str(int(x))

def filtro_checklist(label, opciones, key, fmt=None):
    disp = fmt if fmt else str
    def _sync_todos():
        val  = st.session_state[f"_ck_{key}_all"]
        srch = st.session_state.get(f"_ck_{key}_srch", "")
        for op in opciones:
            if not srch or srch.lower() in disp(op).lower():
                st.session_state[f"_ck_{key}_{op}"] = val
    if f"_ck_{key}_all" not in st.session_state:
        st.session_state[f"_ck_{key}_all"] = True
        for op in opciones: st.session_state[f"_ck_{key}_{op}"] = True
    n_sel = sum(1 for op in opciones if st.session_state.get(f"_ck_{key}_{op}", True))
    n_tot = len(opciones)
    btn_lbl = (f"{label}" if n_sel == n_tot
               else (f"{label} ⚠ Ninguno" if n_sel == 0
               else f"{label} 🔵 {n_sel}/{n_tot}"))
    with st.popover(btn_lbl, use_container_width=True):
        buscar = st.text_input("", key=f"_ck_{key}_srch",
                               placeholder="🔍 Buscar...", label_visibility="collapsed")
        vis = [o for o in opciones if not buscar or buscar.lower() in disp(o).lower()]
        all_checked = all(st.session_state.get(f"_ck_{key}_{op}", True) for op in opciones)
        try:
            st.session_state[f"_ck_{key}_all"] = all_checked
        except Exception:
            pass
        st.checkbox("(Seleccionar todo)", key=f"_ck_{key}_all", on_change=_sync_todos)
        with st.container(height=min(220, max(90, len(vis) * 28))):
            for op in vis:
                ik = f"_ck_{key}_{op}"
                if ik not in st.session_state: st.session_state[ik] = True
                st.checkbox(disp(op), key=ik)
    return [op for op in opciones if st.session_state.get(f"_ck_{key}_{op}", True)]

def color_rec_oc(val):
    if val in ('Recepción Completa','Recepción Cerrada'):
        return f"color:{C_OK};font-weight:bold"
    if val in ('Recepción Parcial',):
        return f"color:{C_ALERTA};font-weight:bold"
    if val in ('Sin Recepciones',):
        return f"color:{C_CRITICO};font-weight:bold"
    return ""

def color_fac(val):
    if str(val).strip() in ('Aprobada','Totalmente Asociada','Ingresado'):
        return f"color:{C_OK};font-weight:bold"
    return ""

def color_desv(val):
    if pd.isna(val): return ""
    if val < -umbral:     return f"color:{C_CRITICO};font-weight:bold"
    if val < -umbral / 2: return f"color:{C_ALERTA};font-weight:bold"
    if val >= 0:          return f"color:{C_OK};font-weight:bold"
    return ""

def color_estado(val):
    if val == "Abierta": return f"color:{C_ALERTA};font-weight:bold"
    if val == "Cerrada": return f"color:{C_OK};font-weight:bold"
    return ""

def color_alerta_fac(val):
    return f"color:{C_CRITICO};font-weight:bold" if val == "FAC > REC" else ""

# ── HTML HELPERS ──────────────────────────────────────────────────────────────
def seccion(icon, titulo, color="#38BDF8", badge_txt=None, badge_color=None):
    bc = badge_color or color
    badge = (f'<span class="sh-badge" style="background:rgba(56,189,248,0.10);color:{bc};">'
             f'{badge_txt}</span>') if badge_txt else ""
    st.markdown(
        f'<div class="sh">'
        f'<div class="sh-bar" style="background:linear-gradient(180deg,{color},{color}44);box-shadow:0 0 10px {color}99;"></div>'
        f'<div class="sh-icon" style="background:linear-gradient(135deg,{color}22,{color}06);border:1px solid {color}38;box-shadow:0 0 12px {color}22;">{icon}</div>'
        f'<div class="sh-txt">{titulo}</div>{badge}</div>',
        unsafe_allow_html=True)

def kpi_card(label, value, delta=None, delta_pos=None):
    if delta is None:        d = ""
    elif delta_pos is True:  d = f'<div class="dp">{delta}</div>'
    elif delta_pos is False: d = f'<div class="dn">{delta}</div>'
    else:                    d = f'<div class="d0">{delta}</div>'
    sz = "kpi-vls" if len(str(value)) > 10 else "kpi-vl"
    return (f'<div class="kpi-c"><div class="kpi-lb">{label}</div>'
            f'<div class="{sz}">{value}</div>{d}</div>')

def kpi_row(items, cols=4):
    html = "".join(kpi_card(*i) for i in items)
    st.markdown(
        f'<div class="kpi-g" style="grid-template-columns:repeat({cols},1fr);">{html}</div>',
        unsafe_allow_html=True)

def mini_kpi(label, value, sub=None, color="#38BDF8"):
    sub_html = (f'<div style="font-size:0.65rem;color:{color};font-weight:600;margin-top:4px;opacity:0.78;">{sub}</div>'
                if sub else "")
    return (
        f'<div style="background:linear-gradient(135deg,{color}18,{color}05);'
        f'border:1px solid {color}38;border-radius:14px;padding:14px 15px 12px;'
        f'min-height:88px;display:flex;flex-direction:column;justify-content:center;'
        f'transition:border-color .2s,box-shadow .2s;">'
        f'<div style="font-size:0.55rem;font-weight:800;letter-spacing:0.12em;'
        f'text-transform:uppercase;color:#475569;margin-bottom:5px;">{label}</div>'
        f'<div style="font-size:1.32rem;font-weight:900;color:{color};line-height:1;'
        f'font-variant-numeric:tabular-nums;text-shadow:0 0 16px {color}66;">{value}</div>'
        f'{sub_html}</div>'
    )

def ring_kpi(label, pct, value_str, sub=None, color="#5470c6"):
    """SVG animated progress ring KPI card."""
    clamped = max(0.0, min(100.0, float(pct or 0)))
    r = 36; cx = cy = 46; circ = 226.2
    offset = circ * (1 - clamped / 100)
    sub_html = f'<div class="ring-sub">{sub}</div>' if sub else ""
    return (
        f'<div class="ring-wrap">'
        f'<div class="ring-lbl">{label}</div>'
        f'<svg width="92" height="92" viewBox="0 0 92 92">'
        f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="rgba(255,255,255,0.07)" stroke-width="5.5"/>'
        f'<circle class="progress" cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="{color}"'
        f' stroke-width="5.5" stroke-linecap="round"'
        f' stroke-dasharray="{circ:.1f}" stroke-dashoffset="{offset:.1f}"'
        f' transform="rotate(-90 {cx} {cy})"/>'
        f'<text x="{cx}" y="{cy+5}" text-anchor="middle" font-size="12.5" font-weight="800"'
        f' fill="{color}" font-family="system-ui,sans-serif">{clamped:.0f}%</text>'
        f'</svg>'
        f'<div class="ring-val" style="color:{color};">{value_str}</div>'
        f'{sub_html}'
        f'</div>'
    )

def ring_row(items, cols=4):
    html = "".join(ring_kpi(*i) for i in items)
    st.markdown(
        f'<div class="kpi-g" style="grid-template-columns:repeat({cols},1fr);">{html}</div>',
        unsafe_allow_html=True)

# ── CARGA DE DATOS ────────────────────────────────────────────────────────────
@st.cache_data
def cargar_presupuesto(mtime):
    cols = ["CC","Especialidad","Capitulo","Capitulo2","Partida","Destino","Partida2",
            "Codigo","Tipo","Ud","Descripcion","CanPres","PU","ImpPres",
            "Rep1","Rep2","CantPartida","Cantidad","Total",
            "DisponibleMat","CostoMat","MargenMat","NOC",
            "DisponibleSub","CostoSub","MargenSub","NContrato","X1","X2","X3","X4"]
    df = pd.read_excel(ARCHIVO_PPTO, sheet_name="APU_PPTO", header=None, skiprows=3, names=cols)
    df = df.dropna(subset=["Descripcion"])
    df = df[df["Tipo"].astype(str).str.strip() == "Material"].copy()
    df = df[df["Ud"].astype(str).str.upper().str.strip() == "M3"].copy()
    df = df[df["Descripcion"].astype(str).apply(es_arido)].copy()
    df["Material"] = df["Descripcion"].astype(str).apply(normalizar)
    df["CC"]       = df["CC"].astype(str).str.strip()
    df["Cantidad"] = pd.to_numeric(df["Cantidad"], errors="coerce").fillna(0)
    df["PU"]       = pd.to_numeric(df["PU"],       errors="coerce").fillna(0)
    df["Total"]    = pd.to_numeric(df["Total"],     errors="coerce").fillna(0)
    df = df[df["CC"].str.match(r"^[A-Z]\d+(\.\d+)?$", na=False)]
    return df

@st.cache_data
def cargar_cc_nombres(mtime):
    df = pd.read_excel(ARCHIVO_PPTO, sheet_name="Nombre_Cuenta_Costo")
    df = df.iloc[:, :2].copy(); df.columns = ["CC","Nombre"]
    df["CC"] = df["CC"].astype(str).str.strip()
    df = df[df["CC"].str.match(r"^[A-Z]\d+(\.\d+)?$", na=False)]
    return dict(zip(df["CC"], df["Nombre"].astype(str).str.strip()))

@st.cache_data
def cargar_recfac(mtime):
    df = pd.read_excel(ARCHIVO_RECFAC, sheet_name="Hoja")
    df.columns = ['N_OC','Nombre_OC','Proveedor','RUT','Fecha_Envio','Moneda','Total_OC',
                  'Forma_Pago','N_Doc_Rec','Monto_Rec','Fecha_Rec','Devolucion',
                  'Saldo_x_Rec','Estado_Guia','N_Factura','Monto_Factura','Folio',
                  'Fecha_Rec_Fac','Estado_Doc_Fac','Estado_Asoc_Fac','Cotizada',
                  'Archivos','Estado_Rec_OC','Razon_Social']
    df['N_OC']          = df['N_OC'].astype(str).str.strip()
    df['Total_OC']      = pd.to_numeric(df['Total_OC'],    errors='coerce').fillna(0)
    df['Monto_Rec']     = pd.to_numeric(df['Monto_Rec'],   errors='coerce').fillna(0)
    df['Saldo_x_Rec']   = pd.to_numeric(df['Saldo_x_Rec'], errors='coerce').fillna(0)
    df['N_Doc_Rec']     = pd.to_numeric(df['N_Doc_Rec'],   errors='coerce').fillna(0)
    df['N_Factura']     = pd.to_numeric(df['N_Factura'],   errors='coerce').fillna(0)
    df['Monto_Factura'] = (df['Monto_Factura'].astype(str)
                           .str.replace('.','',regex=False)
                           .str.replace(',','.',regex=False))
    df['Monto_Factura'] = pd.to_numeric(df['Monto_Factura'], errors='coerce').fillna(0)
    df['Fecha_Rec']     = pd.to_datetime(df['Fecha_Rec'],     errors='coerce')
    df['Fecha_Rec_Fac'] = pd.to_datetime(df['Fecha_Rec_Fac'], errors='coerce')
    df['Fecha_Envio']   = pd.to_datetime(df['Fecha_Envio'],   errors='coerce')
    return df

@st.cache_data
def cargar_nc(mtime):
    """Carga notas de crédito/corrección y devuelve mapa {N_Factura → Monto_NC_total}."""
    try:
        df = pd.read_excel(ARCHIVO_NC, sheet_name="Documentos_Notas_Correccion", header=7)
        df["_fac"] = pd.to_numeric(
            df["Facturas Asociadas a NC"].astype(str).str.strip(), errors="coerce")
        df["Monto_NC"] = pd.to_numeric(
            df["Monto Total"].astype(str)
              .str.replace(".", "", regex=False)
              .str.replace(",", ".", regex=False),
            errors="coerce").fillna(0)
        return df.dropna(subset=["_fac"]).groupby("_fac")["Monto_NC"].sum()
    except Exception:
        return pd.Series(dtype=float)

@st.cache_data
def cargar_oc(mtime):
    df = pd.read_excel(ARCHIVO_OC, sheet_name=0, header=None, skiprows=12)
    df.columns = ["N_OC","Nombre_OC","Fecha","Fecha_Despacho","Metodo_Despacho","Obra",
                  "Razon_Social","Rut","Proveedor","Cod_Maestro","Descripcion","Glosa",
                  "CC","Cuenta_Costo","Partida","Unidad","Cantidad","Moneda",
                  "PU","PU_Desc","Descuento","SubTotal","Cant_Recibida","Monto_Recibido",
                  "Devolucion","Saldo_Recibir","Facturado","Monto_Cerrado","Estado"]
    df = df.dropna(subset=["Descripcion"])
    df = df[df["Descripcion"].astype(str).apply(es_arido)].copy()
    df["Material"]       = df["Descripcion"].astype(str).apply(normalizar)
    df["CC"]             = df["CC"].astype(str).str.strip()
    df["Cantidad"]       = pd.to_numeric(df["Cantidad"],       errors="coerce").fillna(0)
    df["SubTotal"]       = pd.to_numeric(df["SubTotal"],       errors="coerce").fillna(0)
    df["Cant_Recibida"]  = pd.to_numeric(df["Cant_Recibida"],  errors="coerce").fillna(0)
    df["Monto_Recibido"] = pd.to_numeric(df["Monto_Recibido"], errors="coerce").fillna(0)
    df["Fecha"]          = pd.to_datetime(df["Fecha"],         errors="coerce")
    cerrada = df["Estado"] == "Cerrada"
    df["Cant_Efectiva"]  = df["Cant_Recibida"].where(cerrada, df["Cantidad"])
    df["Monto_Efectivo"] = df["Monto_Recibido"].where(cerrada, df["SubTotal"])
    return df

try:
    mtime_ppto = ARCHIVO_PPTO.stat().st_mtime
    mtime_oc   = ARCHIVO_OC.stat().st_mtime
    df_ppto    = cargar_presupuesto(mtime_ppto)
    df_oc      = cargar_oc(mtime_oc)
    cc_map     = cargar_cc_nombres(mtime_ppto)
except FileNotFoundError as e:
    st.error(f"Archivo no encontrado: {e}"); st.stop()

try:
    mtime_rf  = ARCHIVO_RECFAC.stat().st_mtime
    df_recfac = cargar_recfac(mtime_rf)
except FileNotFoundError:
    df_recfac = pd.DataFrame()

try:
    mtime_nc  = ARCHIVO_NC.stat().st_mtime
    nc_by_fac = cargar_nc(mtime_nc)
except FileNotFoundError:
    nc_by_fac = pd.Series(dtype=float)

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="padding:4px 0 12px;">
      <div style="font-size:0.58rem;font-weight:700;letter-spacing:0.18em;text-transform:uppercase;color:#38BDF8;margin-bottom:3px;">CONSTRUCTORA LONDRES</div>
      <div style="font-size:1.05rem;font-weight:800;color:#F1F5F9;letter-spacing:0.02em;">🏗️ CONTROL DE ÁRIDOS</div>
      <div style="font-size:0.75rem;font-weight:600;color:#475569;letter-spacing:0.04em;margin-top:3px;">{PROYECTO_CODIGO} — {PROYECTO_NOMBRE}</div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()
    st.markdown('<div style="font-size:0.65rem;font-weight:700;letter-spacing:0.13em;color:#475569;text-transform:uppercase;margin-bottom:8px;">FILTROS GLOBALES</div>', unsafe_allow_html=True)
    todas_cc  = sorted(set(df_ppto["CC"].tolist() + df_oc["CC"].tolist()))
    todos_mat = sorted(set(df_ppto["Material"].tolist() + df_oc["Material"].tolist()))
    sel_cc  = filtro_checklist("Cuenta de Costo", todas_cc,  "g_cc",  fmt=lambda x: cc_map.get(x, x))
    sel_mat = filtro_checklist("Tipo de Árido",   todos_mat, "g_mat")
    umbral  = st.slider("Umbral alerta desvío (%)", 0, 50, 10, 5)
    st.divider()
    ultima_mod = datetime.fromtimestamp(max(mtime_ppto, mtime_oc)).strftime("%d/%m/%Y %H:%M")
    st.markdown(f'<div style="font-size:0.70rem;color:#334155;">📁 Datos al: <b style="color:#64748B;">{ultima_mod}</b></div>', unsafe_allow_html=True)
    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    if st.button("🔄 Actualizar datos", use_container_width=True):
        st.cache_data.clear(); st.rerun()
    st.markdown(f'<div style="font-size:0.68rem;color:#1E293B;margin-top:6px;">⏱ Refresco cada {INTERVALO}s</div>', unsafe_allow_html=True)
    if len(_TODOS_PROJS) > 1:
        st.divider()
        st.markdown('<div style="font-size:0.65rem;font-weight:700;letter-spacing:0.13em;color:#475569;text-transform:uppercase;margin-bottom:8px;">🏗️ CAMBIAR PROYECTO</div>', unsafe_allow_html=True)
        _projs_info = []
        for _d in _TODOS_PROJS:
            _c = json.loads((_d / "proyecto.json").read_text(encoding="utf-8"))
            _projs_info.append(_c)
        _proj_labels = [f"{p.get('codigo','')} — {p.get('nombre','')}" for p in _projs_info]
        _current_idx = next((i for i, p in enumerate(_projs_info) if p.get("codigo","") == PROYECTO_CODIGO), 0)
        _sel_idx = st.selectbox("", options=range(len(_projs_info)),
                                format_func=lambda i: _proj_labels[i],
                                index=_current_idx, key="proj_switch",
                                label_visibility="collapsed")
        if _sel_idx != _current_idx:
            st.query_params["proyecto"] = _projs_info[_sel_idx].get("codigo", "")
            st.rerun()

# ── DATOS FILTRADOS ───────────────────────────────────────────────────────────
df_ppto_f = df_ppto[df_ppto["CC"].isin(sel_cc) & df_ppto["Material"].isin(sel_mat)]
df_oc_f   = df_oc[df_oc["CC"].isin(sel_cc)     & df_oc["Material"].isin(sel_mat)]
_oc_aridos = set(df_oc_f["N_OC"].astype(str).str.strip())
df_recfac_f = (df_recfac[df_recfac["N_OC"].isin(_oc_aridos)].copy()
               if not df_recfac.empty else pd.DataFrame())

ppto_agg = df_ppto_f.groupby(["CC","Material"]).agg(
    Cant_Ppto=("Cantidad","sum"), Monto_Ppto=("Total","sum")).reset_index()
oc_agg = df_oc_f.groupby(["CC","Material"]).agg(
    Cant_OC=("Cant_Efectiva","sum"), Monto_OC=("Monto_Efectivo","sum"),
    Cant_Recibida=("Cant_Recibida","sum"), Monto_Recibido=("Monto_Recibido","sum")).reset_index()

resumen = pd.merge(ppto_agg, oc_agg, on=["CC","Material"], how="outer").fillna(0)
resumen = resumen.sort_values(["CC","Material"]).reset_index(drop=True)
resumen["CC_Label"]     = resumen["CC"].map(cc_map).fillna(resumen["CC"])
ppto_nz = resumen["Monto_Ppto"].replace(0, float("nan"))
cant_nz = resumen["Cant_Ppto"].replace(0, float("nan"))
resumen["Desv_Monto_%"] = ((resumen["Monto_Ppto"]-resumen["Monto_OC"])/ppto_nz*100).round(1)
resumen["Desv_Cant_%"]  = ((resumen["Cant_Ppto"] -resumen["Cant_OC"] )/cant_nz *100).round(1)
resumen["Avance_%"]     = (resumen["Cant_Recibida"]/cant_nz*100).round(1)
resumen["Semaforo"]  = resumen["Desv_Monto_%"].apply(lambda x: semaforo(x, umbral))
resumen["Sem_Monto"] = resumen["Desv_Monto_%"].apply(lambda x: semaforo(x, umbral))
resumen["Sem_Cant"]  = resumen["Desv_Cant_%"].apply(lambda x: semaforo(x, umbral))
_sin_ppto_fin  = (resumen["Monto_Ppto"]==0) & (resumen["Monto_OC"]>0)
_sin_ppto_cant = (resumen["Cant_Ppto"] ==0) & (resumen["Cant_OC"] >0)
resumen.loc[_sin_ppto_fin,  ["Semaforo","Sem_Monto"]] = "🔴"
resumen.loc[_sin_ppto_cant, "Sem_Cant"] = "🔴"

global_cc = resumen.groupby("CC").agg(
    Monto_Ppto=("Monto_Ppto","sum"), Monto_OC=("Monto_OC","sum"),
    Monto_Recibido=("Monto_Recibido","sum"),
    Cant_Ppto=("Cant_Ppto","sum"), Cant_OC=("Cant_OC","sum"),
    Cant_Recibida=("Cant_Recibida","sum")).reset_index()
global_cc["CC_Label"] = global_cc["CC"].map(cc_map).fillna(global_cc["CC"])
_pnz = global_cc["Monto_Ppto"].replace(0, float("nan"))
_cnz = global_cc["Cant_Ppto"].replace(0, float("nan"))
global_cc["Desv_Fin_%"]  = ((global_cc["Monto_Ppto"]-global_cc["Monto_OC"])/_pnz*100).round(1)
global_cc["Desv_Cant_%"] = ((global_cc["Cant_Ppto"] -global_cc["Cant_OC"] )/_cnz *100).round(1)
global_cc["Sem_Fin"]  = global_cc["Desv_Fin_%"].apply(lambda x: semaforo(x, umbral))
global_cc["Sem_Cant"] = global_cc["Desv_Cant_%"].apply(lambda x: semaforo(x, umbral))
global_cc.loc[(global_cc["Monto_Ppto"]==0)&(global_cc["Monto_OC"]>0), "Sem_Fin"]  = "🔴"
global_cc.loc[(global_cc["Cant_Ppto"] ==0)&(global_cc["Cant_OC"] >0), "Sem_Cant"] = "🔴"
global_cc = global_cc.sort_values("Desv_Fin_%", ascending=True)

# ── KPIs ──────────────────────────────────────────────────────────────────────
tot_ppto          = resumen["Monto_Ppto"].sum()
tot_oc            = resumen["Monto_OC"].sum()
# Usar RecFac como fuente de verdad para Recibido y OC en sistema
if not df_recfac_f.empty:
    # Recepciones únicas (sin duplicar por doc)
    _rec_docs_kpi = df_recfac_f[df_recfac_f["N_Doc_Rec"] > 0].drop_duplicates(["N_OC","N_Doc_Rec"])
    tot_recibido  = _rec_docs_kpi["Monto_Rec"].sum()

    # Estado por N_OC desde el archivo OC (str para coincidir con RecFac)
    _oc_estado_map = dict(zip(df_oc_f["N_OC"].astype(str).str.strip(), df_oc_f["Estado"]))

    # Suma de recepciones reales por N_OC (para OCs cerradas)
    _rec_por_oc = _rec_docs_kpi.groupby("N_OC")["Monto_Rec"].sum()

    # Una fila por N_OC en RecFac — Total_OC no se repite al sumar
    _oc_unicas = df_recfac_f.drop_duplicates("N_OC")[["N_OC","Total_OC"]].copy()
    _oc_unicas["Estado"] = _oc_unicas["N_OC"].map(_oc_estado_map).fillna("Abierta")

    # Cerrada → monto efectivo = lo realmente recibido (OC ya no generará más recepciones)
    # Abierta → monto comprometido = Total_OC completo
    _oc_unicas["Monto_Ef"] = _oc_unicas.apply(
        lambda r: _rec_por_oc.get(r["N_OC"], r["Total_OC"])
                  if r["Estado"] == "Cerrada" else r["Total_OC"], axis=1)

    _oc_recfac = _oc_unicas["Monto_Ef"].sum()
    if _oc_recfac > 0:
        tot_oc = _oc_recfac
    # Invariante: OC nunca menor a lo recepcionado
    tot_oc = max(tot_oc, tot_recibido)
else:
    tot_recibido  = resumen["Monto_Recibido"].sum()
tot_cant_ppto     = resumen["Cant_Ppto"].sum()
tot_cant_oc       = resumen["Cant_OC"].sum()
tot_cant_recibida = resumen["Cant_Recibida"].sum()
n_activas         = df_oc_f[df_oc_f["Estado"]=="Abierta"]["N_OC"].nunique()
n_alertas         = (resumen["Semaforo"]=="🔴").sum()
desv_global       = ((tot_ppto-tot_oc)/tot_ppto*100) if tot_ppto else 0
ejec_pct          = (tot_recibido/tot_oc*100) if tot_oc else 0
compromiso_pct    = (tot_oc/tot_ppto*100) if tot_ppto else 0
avance_cant_pct   = (tot_cant_recibida/tot_cant_ppto*100) if tot_cant_ppto else 0
desv_cant_global  = ((tot_cant_ppto-tot_cant_oc)/tot_cant_ppto*100) if tot_cant_ppto else 0
n_sin_ppto        = int(_sin_ppto_fin.sum())
n_cc_criticas     = int((global_cc["Sem_Fin"]=="🔴").sum())
_saldo_fin        = tot_ppto - tot_oc

if compromiso_pct > 100:
    _est_txt   = "CRÍTICO";    _est_css = "stat-critico"; _est_color = C_CRITICO
    _est_detail = f"OC supera el presupuesto total · {compromiso_pct:.1f}% comprometido · Exceso MM$ {abs(_saldo_fin)/1e6:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",",")
elif compromiso_pct >= 95:
    _est_txt   = "EN RIESGO";  _est_css = "stat-riesgo";  _est_color = C_ALERTA
    _est_detail = f"Presupuesto casi agotado · {compromiso_pct:.1f}% comprometido · Saldo MM$ {_saldo_fin/1e6:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",",")
else:
    _est_txt   = "CONTROLADO"; _est_css = "stat-normal";  _est_color = C_OK
    _est_detail = f"{compromiso_pct:.1f}% del presupuesto comprometido · Saldo disponible MM$ {_saldo_fin/1e6:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",",")

# ═════════════════════════════════════════════════════════════════════════════
# HEADER — CARTA DE PRESENTACIÓN PREMIUM
# ═════════════════════════════════════════════════════════════════════════════
_dif_cant    = tot_cant_ppto - tot_cant_oc
_ppto_mm     = _n(tot_ppto/1e6,           2, "MM$ ")
_oc_mm       = _n(tot_oc/1e6,             2, "MM$ ")
_rec_mm      = _n(tot_recibido/1e6,       2, "MM$ ")
_sal_mm      = _n(abs(_saldo_fin)/1e6,    2, "MM$ ")
_ppto_m3     = _n(tot_cant_ppto,           0, sfx=" m³")
_oc_m3       = _n(tot_cant_oc,             0, sfx=" m³")
_rec_m3      = _n(tot_cant_recibida,       0, sfx=" m³")
_sal_m3      = _n(abs(_dif_cant),          0, sfx=" m³")
_desv_fmt    = f"{desv_global:+.1f}%".replace(".", ",")
_desv_c_fmt  = f"{desv_cant_global:+.1f}%".replace(".", ",")
_ejec_fmt    = f"{ejec_pct:.1f}%".replace(".", ",")
_avance_fmt  = f"{avance_cant_pct:.1f}%".replace(".", ",")
_comp_fmt    = f"{compromiso_pct:.1f}%".replace(".", ",")
_fecha_str   = datetime.fromtimestamp(max(mtime_ppto, mtime_oc)).strftime("%d/%m/%Y %H:%M")
_hero_rgb    = "255,71,87" if compromiso_pct > 100 else ("255,179,0" if compromiso_pct >= 95 else "35,209,96")
_oc_clr      = C_CRITICO if tot_oc >= tot_ppto else C_OK
_sal_clr     = C_OK if _saldo_fin > 0 else C_CRITICO
_oc_m3_clr   = C_CRITICO if tot_cant_oc >= tot_cant_ppto else C_OK
_sal_m3_clr  = C_OK if _dif_cant > 0 else C_CRITICO
_sal_lbl     = "disponible" if _saldo_fin > 0 else "exceso"
_sal_m3_lbl  = "disponible" if _dif_cant > 0 else "exceso"
_desv_clr    = "#23D160" if desv_global     >= 0 else "#FF4757"
_desv_c_clr  = "#23D160" if desv_cant_global >= 0 else "#FF4757"
_pb_oc_w     = f"{min(100.0, compromiso_pct):.1f}"
_pb_rec_w    = f"{min(100.0, tot_recibido / tot_ppto * 100 if tot_ppto else 0):.1f}"
_pb_oc_m3_w  = f"{min(100.0, tot_cant_oc / tot_cant_ppto * 100 if tot_cant_ppto else 0):.1f}"
_pb_rec_m3_w = f"{min(100.0, tot_cant_recibida / tot_cant_ppto * 100 if tot_cant_ppto else 0):.1f}"
_al_bg       = "rgba(255,71,87,0.13)"  if n_alertas  > 0 else "rgba(100,116,139,0.07)"
_al_brd      = "rgba(255,71,87,0.38)"  if n_alertas  > 0 else "rgba(100,116,139,0.14)"
_al_clr      = "#FF4757"               if n_alertas  > 0 else "#475569"
_sp_bg       = "rgba(255,71,87,0.13)"  if n_sin_ppto > 0 else "rgba(100,116,139,0.07)"
_sp_brd      = "rgba(255,71,87,0.38)"  if n_sin_ppto > 0 else "rgba(100,116,139,0.14)"
_sp_clr      = "#FF4757"               if n_sin_ppto > 0 else "#475569"

st.markdown(f"""
<div style="background:linear-gradient(90deg,rgba(56,189,248,0.08),rgba(56,189,248,0.02) 55%,transparent);
            border:1px solid rgba(56,189,248,0.15);border-radius:18px;padding:16px 22px;margin-bottom:18px;
            display:flex;align-items:center;gap:16px;">
  <div>
    <div style="font-size:0.68rem;font-weight:800;letter-spacing:0.28em;text-transform:uppercase;color:#38BDF8;margin-bottom:4px;text-shadow:0 0 14px rgba(56,189,248,0.40);">CONSTRUCTORA LONDRES</div>
    <div style="font-size:2.0rem;font-weight:900;color:#F8FAFC;letter-spacing:0.02em;line-height:1.05;text-shadow:0 0 24px rgba(255,255,255,0.12);">🏗️ CONTROL DE ÁRIDOS</div>
  </div>
  <div style="width:1px;height:38px;background:rgba(56,189,248,0.20);flex-shrink:0;"></div>
  <div style="background:rgba(56,189,248,0.09);border:1px solid rgba(56,189,248,0.22);border-radius:10px;padding:6px 14px;">
    <div style="font-size:0.55rem;font-weight:700;letter-spacing:0.13em;color:#64748B;">PROYECTO</div>
    <div style="font-size:0.90rem;font-weight:800;color:#F1F5F9;">{PROYECTO_CODIGO} · {PROYECTO_NOMBRE}</div>
  </div>
  <div style="flex:1;"></div>
  <div style="display:flex;gap:9px;align-items:center;flex-wrap:wrap;">
    <div style="background:{_al_bg};border:1px solid {_al_brd};border-radius:8px;padding:5px 12px;font-size:0.70rem;font-weight:700;color:{_al_clr};">⚠ {n_alertas} alertas</div>
    <div style="background:{_sp_bg};border:1px solid {_sp_brd};border-radius:8px;padding:5px 12px;font-size:0.70rem;font-weight:700;color:{_sp_clr};">🔴 {n_sin_ppto} sin ppto</div>
    <div style="background:rgba(255,255,255,0.025);border:1px solid rgba(255,255,255,0.06);border-radius:8px;padding:5px 12px;font-size:0.67rem;color:#334155;">📅 {_fecha_str}</div>
  </div>
</div>
""", unsafe_allow_html=True)

seccion("📊", "RESUMEN EJECUTIVO", color="#38BDF8")

st.markdown(f"""
<div style="display:grid;grid-template-columns:1fr 2.25fr;gap:18px;margin-bottom:18px;">
  <div style="background:linear-gradient(145deg,rgba({_hero_rgb},0.18) 0%,rgba({_hero_rgb},0.04) 100%);
              border:1.5px solid rgba({_hero_rgb},0.50);
              box-shadow:0 0 64px rgba({_hero_rgb},0.15),inset 0 0 40px rgba({_hero_rgb},0.04);
              border-radius:20px;padding:26px 22px;display:flex;flex-direction:column;justify-content:center;gap:12px;">
    <div style="font-size:0.57rem;font-weight:800;letter-spacing:0.22em;text-transform:uppercase;color:rgba({_hero_rgb},0.65);">ESTADO DEL PROYECTO</div>
    <div style="font-size:3.0rem;font-weight:900;line-height:1.0;letter-spacing:0.02em;color:rgb({_hero_rgb});
                text-shadow:0 0 32px rgba({_hero_rgb},0.65),0 0 72px rgba({_hero_rgb},0.28);">{_est_txt}</div>
    <div>
      <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:5px;">
        <span style="font-size:0.59rem;font-weight:700;letter-spacing:0.10em;text-transform:uppercase;color:#475569;">COMPROMISO PRESUPUESTO</span>
        <span style="font-size:0.88rem;font-weight:900;color:rgb({_hero_rgb});text-shadow:0 0 12px rgba({_hero_rgb},0.50);">{_comp_fmt}</span>
      </div>
      <div style="height:6px;background:rgba(255,255,255,0.055);border-radius:99px;overflow:hidden;">
        <div style="height:100%;width:{_pb_oc_w}%;border-radius:99px;
                    background:linear-gradient(90deg,rgb({_hero_rgb}),rgba({_hero_rgb},0.38));
                    box-shadow:0 0 12px rgba({_hero_rgb},0.60);"></div>
      </div>
    </div>
    <div style="font-size:0.71rem;color:rgba(255,255,255,0.40);line-height:1.55;border-top:1px solid rgba({_hero_rgb},0.14);padding-top:10px;">{_est_detail}</div>
    <div style="display:flex;gap:20px;">
      <div style="text-align:center;">
        <div style="font-size:1.15rem;font-weight:800;color:#38BDF8;text-shadow:0 0 12px rgba(56,189,248,0.40);">{n_activas}</div>
        <div style="font-size:0.55rem;font-weight:700;color:#334155;letter-spacing:0.10em;text-transform:uppercase;">OC activas</div>
      </div>
      <div style="text-align:center;">
        <div style="font-size:1.15rem;font-weight:800;color:{_al_clr};text-shadow:0 0 10px {_al_clr}44;">{n_alertas}</div>
        <div style="font-size:0.55rem;font-weight:700;color:#334155;letter-spacing:0.10em;text-transform:uppercase;">alertas</div>
      </div>
      <div style="text-align:center;">
        <div style="font-size:1.15rem;font-weight:800;color:{_sp_clr};text-shadow:0 0 10px {_sp_clr}44;">{n_sin_ppto}</div>
        <div style="font-size:0.55rem;font-weight:700;color:#334155;letter-spacing:0.10em;text-transform:uppercase;">sin ppto</div>
      </div>
    </div>
  </div>
  <div style="display:flex;flex-direction:column;gap:14px;">
    <div style="display:flex;align-items:center;gap:9px;">
      <div style="width:18px;height:2.5px;background:#38BDF8;border-radius:2px;box-shadow:0 0 8px #38BDF8;flex-shrink:0;"></div>
      <div style="font-size:0.62rem;font-weight:800;letter-spacing:0.18em;text-transform:uppercase;color:#38BDF8;">💰 FINANCIERO (MM$)</div>
    </div>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;">
      <div style="background:linear-gradient(135deg,rgba(84,112,198,0.14),rgba(84,112,198,0.03));border:1px solid rgba(84,112,198,0.30);border-radius:14px;padding:14px 14px 12px;display:flex;flex-direction:column;gap:3px;">
        <div style="font-size:0.55rem;font-weight:800;letter-spacing:0.12em;text-transform:uppercase;color:#475569;">PRESUPUESTO</div>
        <div style="font-size:1.32rem;font-weight:900;color:#7C9FE4;line-height:1;font-variant-numeric:tabular-nums;text-shadow:0 0 16px rgba(84,112,198,0.42);">{_ppto_mm}</div>
        <div style="font-size:0.62rem;color:#334155;">referencia base</div>
        <div style="margin-top:8px;height:3px;background:rgba(255,255,255,0.05);border-radius:99px;overflow:hidden;"><div style="height:100%;width:100%;background:linear-gradient(90deg,#5470c6,rgba(84,112,198,0.38));border-radius:99px;box-shadow:0 0 6px rgba(84,112,198,0.50);"></div></div>
      </div>
      <div style="background:linear-gradient(135deg,{_oc_clr}22,{_oc_clr}06);border:1px solid {_oc_clr}44;border-radius:14px;padding:14px 14px 12px;display:flex;flex-direction:column;gap:3px;">
        <div style="font-size:0.55rem;font-weight:800;letter-spacing:0.12em;text-transform:uppercase;color:#475569;">OC EFECTIVO</div>
        <div style="font-size:1.32rem;font-weight:900;color:{_oc_clr};line-height:1;font-variant-numeric:tabular-nums;text-shadow:0 0 16px {_oc_clr}88;">{_oc_mm}</div>
        <div style="font-size:0.62rem;color:{_desv_clr};font-weight:600;">{_desv_fmt} vs ppto</div>
        <div style="margin-top:8px;height:3px;background:rgba(255,255,255,0.05);border-radius:99px;overflow:hidden;"><div style="height:100%;width:{_pb_oc_w}%;background:linear-gradient(90deg,{_oc_clr},{_oc_clr}55);border-radius:99px;box-shadow:0 0 6px {_oc_clr};"></div></div>
      </div>
      <div style="background:linear-gradient(135deg,rgba(56,189,248,0.14),rgba(56,189,248,0.03));border:1px solid rgba(56,189,248,0.26);border-radius:14px;padding:14px 14px 12px;display:flex;flex-direction:column;gap:3px;">
        <div style="font-size:0.55rem;font-weight:800;letter-spacing:0.12em;text-transform:uppercase;color:#475569;">RECIBIDO</div>
        <div style="font-size:1.32rem;font-weight:900;color:#38BDF8;line-height:1;font-variant-numeric:tabular-nums;text-shadow:0 0 16px rgba(56,189,248,0.52);">{_rec_mm}</div>
        <div style="font-size:0.62rem;color:#38BDF8;font-weight:600;">{_ejec_fmt} ejecutado</div>
        <div style="margin-top:8px;height:3px;background:rgba(255,255,255,0.05);border-radius:99px;overflow:hidden;"><div style="height:100%;width:{_pb_rec_w}%;background:linear-gradient(90deg,#38BDF8,rgba(56,189,248,0.48));border-radius:99px;box-shadow:0 0 6px rgba(56,189,248,0.58);"></div></div>
      </div>
      <div style="background:linear-gradient(135deg,{_sal_clr}22,{_sal_clr}06);border:1px solid {_sal_clr}40;border-radius:14px;padding:14px 14px 12px;display:flex;flex-direction:column;gap:3px;">
        <div style="font-size:0.55rem;font-weight:800;letter-spacing:0.12em;text-transform:uppercase;color:#475569;">SALDO PPTO</div>
        <div style="font-size:1.32rem;font-weight:900;color:{_sal_clr};line-height:1;font-variant-numeric:tabular-nums;text-shadow:0 0 16px {_sal_clr}88;">{_sal_mm}</div>
        <div style="font-size:0.62rem;color:{_sal_clr};font-weight:600;">{_sal_lbl}</div>
        <div style="margin-top:8px;height:3px;background:rgba(255,255,255,0.05);border-radius:99px;overflow:hidden;"></div>
      </div>
    </div>
    <div style="display:flex;align-items:center;gap:9px;">
      <div style="width:18px;height:2.5px;background:#A78BFA;border-radius:2px;box-shadow:0 0 8px #A78BFA;flex-shrink:0;"></div>
      <div style="font-size:0.62rem;font-weight:800;letter-spacing:0.18em;text-transform:uppercase;color:#A78BFA;">📦 CANTIDADES (m³)</div>
    </div>
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;">
      <div style="background:linear-gradient(135deg,rgba(84,112,198,0.14),rgba(84,112,198,0.03));border:1px solid rgba(84,112,198,0.30);border-radius:14px;padding:14px 14px 12px;display:flex;flex-direction:column;gap:3px;">
        <div style="font-size:0.55rem;font-weight:800;letter-spacing:0.12em;text-transform:uppercase;color:#475569;">PRESUPUESTO</div>
        <div style="font-size:1.25rem;font-weight:900;color:#7C9FE4;line-height:1;font-variant-numeric:tabular-nums;text-shadow:0 0 14px rgba(84,112,198,0.38);">{_ppto_m3}</div>
        <div style="font-size:0.62rem;color:#334155;">referencia base</div>
        <div style="margin-top:8px;height:3px;background:rgba(255,255,255,0.05);border-radius:99px;overflow:hidden;"><div style="height:100%;width:100%;background:linear-gradient(90deg,#5470c6,rgba(84,112,198,0.38));border-radius:99px;box-shadow:0 0 6px rgba(84,112,198,0.50);"></div></div>
      </div>
      <div style="background:linear-gradient(135deg,{_oc_m3_clr}22,{_oc_m3_clr}06);border:1px solid {_oc_m3_clr}44;border-radius:14px;padding:14px 14px 12px;display:flex;flex-direction:column;gap:3px;">
        <div style="font-size:0.55rem;font-weight:800;letter-spacing:0.12em;text-transform:uppercase;color:#475569;">OC EFECTIVO</div>
        <div style="font-size:1.25rem;font-weight:900;color:{_oc_m3_clr};line-height:1;font-variant-numeric:tabular-nums;text-shadow:0 0 14px {_oc_m3_clr}88;">{_oc_m3}</div>
        <div style="font-size:0.62rem;color:{_desv_c_clr};font-weight:600;">{_desv_c_fmt} vs ppto</div>
        <div style="margin-top:8px;height:3px;background:rgba(255,255,255,0.05);border-radius:99px;overflow:hidden;"><div style="height:100%;width:{_pb_oc_m3_w}%;background:linear-gradient(90deg,{_oc_m3_clr},{_oc_m3_clr}55);border-radius:99px;box-shadow:0 0 6px {_oc_m3_clr};"></div></div>
      </div>
      <div style="background:linear-gradient(135deg,rgba(167,139,250,0.14),rgba(167,139,250,0.03));border:1px solid rgba(167,139,250,0.25);border-radius:14px;padding:14px 14px 12px;display:flex;flex-direction:column;gap:3px;">
        <div style="font-size:0.55rem;font-weight:800;letter-spacing:0.12em;text-transform:uppercase;color:#475569;">RECIBIDO</div>
        <div style="font-size:1.25rem;font-weight:900;color:#A78BFA;line-height:1;font-variant-numeric:tabular-nums;text-shadow:0 0 14px rgba(167,139,250,0.48);">{_rec_m3}</div>
        <div style="font-size:0.62rem;color:#A78BFA;font-weight:600;">{_avance_fmt} recepcionado</div>
        <div style="margin-top:8px;height:3px;background:rgba(255,255,255,0.05);border-radius:99px;overflow:hidden;"><div style="height:100%;width:{_pb_rec_m3_w}%;background:linear-gradient(90deg,#A78BFA,rgba(167,139,250,0.48));border-radius:99px;box-shadow:0 0 6px rgba(167,139,250,0.52);"></div></div>
      </div>
      <div style="background:linear-gradient(135deg,{_sal_m3_clr}22,{_sal_m3_clr}06);border:1px solid {_sal_m3_clr}40;border-radius:14px;padding:14px 14px 12px;display:flex;flex-direction:column;gap:3px;">
        <div style="font-size:0.55rem;font-weight:800;letter-spacing:0.12em;text-transform:uppercase;color:#475569;">SALDO</div>
        <div style="font-size:1.25rem;font-weight:900;color:{_sal_m3_clr};line-height:1;font-variant-numeric:tabular-nums;text-shadow:0 0 14px {_sal_m3_clr}88;">{_sal_m3}</div>
        <div style="font-size:0.62rem;color:{_sal_m3_clr};font-weight:600;">{_sal_m3_lbl}</div>
        <div style="margin-top:8px;height:3px;background:rgba(255,255,255,0.05);border-radius:99px;overflow:hidden;"></div>
      </div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

_n_res = max(len(resumen), 1)
st.divider()

# ═════════════════════════════════════════════════════════════════════════════
# DIAGNÓSTICO DE RIESGO
# ═════════════════════════════════════════════════════════════════════════════
seccion("🎯", "DIAGNÓSTICO DE RIESGO", color="#FF4757",
        badge_txt=f"{n_cc_criticas} CC críticas" if n_cc_criticas else "Sin alertas CC",
        badge_color="#FF4757" if n_cc_criticas else "#23D160")

# ── Hero Risk Summary ─────────────────────────────────────────────────────────
_gc_exc = global_cc[global_cc["Monto_OC"] > global_cc["Monto_Ppto"]]
_rsk_total_exp = _gc_exc["Monto_OC"].sum() / 1e6
_rsk_total_exp_fmt = f"MM$ {_rsk_total_exp:,.1f}".replace(".", "\x00").replace(",", ".").replace("\x00", ",")
_rsk_worst_cc  = global_cc.sort_values("Desv_Fin_%").iloc[0] if not global_cc.empty else None
_rsk_score     = max(0, min(100, int(100 - compromiso_pct))) if tot_ppto else 50
_score_color   = "#23D160" if _rsk_score >= 15 else ("#FFB300" if _rsk_score >= 5 else "#FF4757")
_score_label   = "SANO" if _rsk_score >= 15 else ("ALERTA" if _rsk_score >= 5 else "CRÍTICO")
_worst_lbl     = (_rsk_worst_cc["CC_Label"] if _rsk_worst_cc is not None else "—")
_worst_desv    = (_rsk_worst_cc["Desv_Fin_%"] if _rsk_worst_cc is not None and not pd.isna(_rsk_worst_cc["Desv_Fin_%"]) else 0)
_score_rgb     = "35,209,96" if _rsk_score >= 15 else ("255,179,0" if _rsk_score >= 5 else "255,71,87")

st.markdown(f"""
<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:18px;">
  <div style="background:linear-gradient(135deg,rgba(255,71,87,0.18) 0%,rgba(255,71,87,0.05) 100%);
              border:1.5px solid rgba(255,71,87,0.55);border-radius:14px;padding:14px 16px 12px;
              box-shadow:0 0 24px rgba(255,71,87,0.20),inset 0 1px 0 rgba(255,255,255,0.05);">
    <div style="font-size:0.55rem;font-weight:800;letter-spacing:0.14em;text-transform:uppercase;color:rgba(255,71,87,0.85);margin-bottom:6px;">CC Críticas</div>
    <div style="font-size:1.55rem;font-weight:900;color:#FF4757;line-height:1;text-shadow:0 0 16px rgba(255,71,87,0.55);">{n_cc_criticas}</div>
    <div style="font-size:0.62rem;color:#64748B;margin-top:4px;">cuentas sobre presupuesto</div>
  </div>
  <div style="background:linear-gradient(135deg,rgba(255,179,0,0.15) 0%,rgba(255,179,0,0.04) 100%);
              border:1.5px solid rgba(255,179,0,0.50);border-radius:14px;padding:14px 16px 12px;
              box-shadow:0 0 24px rgba(255,179,0,0.16),inset 0 1px 0 rgba(255,255,255,0.05);">
    <div style="font-size:0.55rem;font-weight:800;letter-spacing:0.14em;text-transform:uppercase;color:rgba(255,179,0,0.90);margin-bottom:6px;">Exposición Total</div>
    <div style="font-size:1.42rem;font-weight:900;color:#FFB300;line-height:1;text-shadow:0 0 16px rgba(255,179,0,0.48);">{_rsk_total_exp_fmt}</div>
    <div style="font-size:0.62rem;color:#64748B;margin-top:4px;">monto en zona de exceso</div>
  </div>
  <div style="background:linear-gradient(135deg,rgba(255,71,87,0.10) 0%,rgba(15,23,42,0.0) 100%);
              border:1px solid rgba(255,255,255,0.09);border-radius:14px;padding:14px 16px 12px;">
    <div style="font-size:0.55rem;font-weight:800;letter-spacing:0.14em;text-transform:uppercase;color:#64748B;margin-bottom:6px;">Mayor Desvío</div>
    <div style="font-size:0.95rem;font-weight:800;color:#FF4757;line-height:1.25;text-shadow:0 0 12px rgba(255,71,87,0.38);">{_worst_lbl}</div>
    <div style="font-size:0.70rem;font-weight:700;color:#FF4757;margin-top:4px;">{_worst_desv:+.1f}% vs ppto</div>
  </div>
</div>
""", unsafe_allow_html=True)

tab_matrix, tab_pareto = st.tabs(["  🔴  Matriz de Riesgo  ","  📊  Pareto de Desvíos  "])

with tab_matrix:
    st.markdown("""
    <div style="background:linear-gradient(90deg,rgba(255,71,87,0.08),rgba(84,112,198,0.06),rgba(35,209,96,0.06));
                border:1px solid rgba(255,255,255,0.07);border-radius:10px;
                padding:10px 16px;margin-bottom:14px;font-size:0.75rem;color:#64748B;line-height:1.6;">
      <span style="color:#FF4757;font-weight:700;">■ Eje X</span> Desviación financiera % ·
      <span style="color:#FFB300;font-weight:700;">■ Eje Y</span> Monto comprometido MM$ ·
      <span style="color:#94A3B8;font-weight:700;">⬤ Tamaño</span> Monto OC ·
      Zona <span style="color:#FF4757;font-weight:700;">izquierda</span> = alta exposición ·
      Zona <span style="color:#23D160;font-weight:700;">derecha</span> = margen disponible
    </div>""", unsafe_allow_html=True)

    gc = global_cc.copy()
    gc["Desv_plot"] = gc["Desv_Fin_%"].fillna(-100)
    gc["OC_MM"]     = gc["Monto_OC"] / 1e6
    gc["Ppto_MM"]   = gc["Monto_Ppto"] / 1e6
    _maxoc          = gc["OC_MM"].max() if gc["OC_MM"].max() > 0 else 1
    gc["size_px"]   = gc["OC_MM"].apply(lambda v: 20 + (v/_maxoc)*55 if v > 0 else 14)
    gc["cat"]       = gc["Sem_Fin"].map({"🔴":"Crítico","🟡":"Alerta","🟢":"Normal","⚪":"Sin ppto"}).fillna("Sin ppto")

    gc = gc.sort_values(["Desv_plot","OC_MM"]).reset_index(drop=True)
    gc["_jx"] = gc["Desv_plot"].astype(float)
    gc["_jy"] = gc["OC_MM"].astype(float)
    _x_tol = max((gc["Desv_plot"].max() - gc["Desv_plot"].min()) * 0.06, 4.0)
    _y_tol = max((gc["OC_MM"].max() - gc["OC_MM"].min()) * 0.08, 0.15)
    gc["_xgrp"] = (gc["Desv_plot"] / _x_tol).round().astype(int)
    gc["_ygrp"] = (gc["OC_MM"]    / _y_tol).round().astype(int)
    gc["_rank"] = gc.groupby(["_xgrp","_ygrp"]).cumcount()
    _ox = [0, _x_tol*0.65, 0,           _x_tol*0.65]
    _oy = [0, 0,           _y_tol*0.95, _y_tol*0.95]
    gc["_jx"] += gc["_rank"].apply(lambda r: _ox[r % 4])
    gc["_jy"] += gc["_rank"].apply(lambda r: _oy[r % 4])

    _ymax    = max(gc["_jy"].max() * 1.22, 1)
    _ymax_v  = round(float(_ymax), 2)
    _xmin_v  = round(float(gc["_jx"].min() - 6) if not gc.empty else -20, 1)
    _xmax_v  = round(float(gc["_jx"].max() + 8) if not gc.empty else 40, 1)

    _RISK_COLORS = {
        "Crítico":  ("#FF4757", "rgba(255,71,87,0.90)",  "rgba(255,71,87,0.35)",  20),
        "Alerta":   ("#FFB300", "rgba(255,179,0,0.90)",  "rgba(255,179,0,0.30)",  12),
        "Normal":   ("#23D160", "rgba(35,209,96,0.90)",  "rgba(35,209,96,0.25)",   8),
        "Sin ppto": ("#64748B", "rgba(100,116,139,0.80)","rgba(100,116,139,0.20)", 6),
    }
    ec_series_mx = []
    for cat_name, (c_hex, c_fill, c_shadow, blur) in _RISK_COLORS.items():
        sub = gc[gc["cat"] == cat_name]
        if sub.empty: continue
        pts = []
        for _, r in sub.iterrows():
            pts.append({
                "value": [round(float(r["_jx"]),3), round(float(r["_jy"]),3), round(float(r["size_px"]),1)],
                "name": r["CC"],
                "ccLabel": r["CC_Label"],
                "desv": round(float(r["Desv_plot"]),1),
                "oc": round(float(r["OC_MM"]),2),
                "ppto": round(float(r["Ppto_MM"]),2),
                "desvCant": round(float(r["Desv_Cant_%"]) if not pd.isna(r["Desv_Cant_%"]) else 0, 1),
            })
        ec_series_mx.append({
            "name": cat_name,
            "type": "scatter",
            "data": pts,
            "symbolSize": "__SYMSIZE__",
            "itemStyle": {
                "color": c_fill,
                "borderColor": c_hex,
                "borderWidth": 2.5,
                "shadowBlur": blur,
                "shadowColor": c_shadow,
            },
            "label": {
                "show": True, "formatter": "{b}",
                "color": "#F1F5F9", "fontSize": 10, "fontWeight": "700",
                "distance": 8,
                "textShadowBlur": 4, "textShadowColor": "rgba(0,0,0,0.8)",
            },
            "emphasis": {
                "scale": 1.18,
                "itemStyle": {"shadowBlur": blur * 2, "shadowColor": c_shadow},
            },
        })

    js_mx = (
        '{"backgroundColor":"transparent","animation":true,"animationDuration":900,"animationEasing":"cubicOut",'
        '"tooltip":{"trigger":"item",'
        '"backgroundColor":"rgba(10,15,30,0.97)",'
        '"borderColor":"rgba(255,255,255,0.12)","borderWidth":1,'
        '"padding":[12,16],'
        '"textStyle":{"color":"#F1F5F9","fontSize":12},'
        '"extraCssText":"border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,0.6);",'
        '"formatter":function(p){'
        'var d=p.data,col=p.color;'
        'var icon=d.desv<0?"🔴 EXCESO":"🟢 MARGEN";'
        'return \'<div style="font-size:13px;font-weight:800;margin-bottom:6px;color:\'+col+\'">\'+d.name+\'</div>\''
        '+\'<div style="color:#94A3B8;font-size:11px;margin-bottom:8px;">\'+d.ccLabel+\'</div>\''
        '+\'<div style="border-top:1px solid rgba(255,255,255,0.08);padding-top:8px;">\''
        '+\'<span style="color:#64748B;">Desv. financiera: </span><b style="color:\'+col+\'">\'+d.desv.toFixed(1)+\'%</b> \'+icon+\'<br/>\''
        '+\'<span style="color:#64748B;">OC comprometido: </span><b>MM$ \'+d.oc.toFixed(2)+\'</b><br/>\''
        '+\'<span style="color:#64748B;">Presupuesto CC: </span><b>MM$ \'+d.ppto.toFixed(2)+\'</b><br/>\''
        '+\'<span style="color:#64748B;">Desv. cantidad: </span><b>\'+d.desvCant.toFixed(1)+\'%</b></div>\';}}'
        ','
        '"toolbox":{"feature":{"saveAsImage":{"title":"Guardar"},"restore":{"title":"Reset"}},'
        '"iconStyle":{"borderColor":"#334155"},"right":12,"top":8},'
        '"legend":{"data":["Crítico","Alerta","Normal","Sin ppto"],'
        '"textStyle":{"color":"#94A3B8","fontSize":11},'
        '"itemStyle":{"borderWidth":0},'
        '"right":90,"top":8,"icon":"circle"},'
        '"xAxis":{"type":"value","name":"Desviación Financiera (%)","nameLocation":"middle","nameGap":36,'
        f'"min":{_xmin_v},"max":{_xmax_v},'
        '"nameTextStyle":{"color":"#475569","fontSize":11,"fontWeight":"600"},'
        '"axisLine":{"show":true,"lineStyle":{"color":"rgba(255,255,255,0.10)","width":1}},'
        '"splitLine":{"lineStyle":{"color":"rgba(255,255,255,0.04)","type":"dashed"}},'
        '"axisLabel":{"color":"#64748B","fontSize":10,"formatter":"{value}%"}},'
        '"yAxis":{"type":"value","name":"Monto Comprometido (MM$)","nameLocation":"middle","nameGap":58,'
        f'"min":0,"max":{_ymax_v},'
        '"nameTextStyle":{"color":"#475569","fontSize":11,"fontWeight":"600"},'
        '"axisLine":{"show":true,"lineStyle":{"color":"rgba(255,255,255,0.10)","width":1}},'
        '"splitLine":{"lineStyle":{"color":"rgba(255,255,255,0.04)","type":"dashed"}},'
        '"axisLabel":{"color":"#64748B","fontSize":10,"formatter":"MM$ {value}"}},'
        '"graphic":['
        '{"type":"rect","shape":{"x":0,"y":0,"width":"50%","height":"100%"},"silent":true,'
        '"style":{"fill":"rgba(255,71,87,0.028)"},"zlevel":-1},'
        '{"type":"rect","shape":{"x":"50%","y":0,"width":"50%","height":"100%"},"silent":true,'
        '"style":{"fill":"rgba(35,209,96,0.022)"},"zlevel":-1},'
        '{"type":"text","left":"62%","top":14,'
        '"style":{"text":"ZONA MARGEN  ✓","fill":"rgba(35,209,96,0.65)","fontSize":10,"fontWeight":"700","letterSpacing":2}},'
        '{"type":"text","left":52,"top":14,'
        '"style":{"text":"◀  ALTA EXPOSICIÓN","fill":"rgba(255,71,87,0.70)","fontSize":10,"fontWeight":"700","letterSpacing":2}},'
        '{"type":"line",'
        '"shape":{"x1":"50%","y1":0,"x2":"50%","y2":"100%"},'
        '"style":{"stroke":"rgba(255,255,255,0.20)","lineDash":[5,5],"lineWidth":1.5},'
        '"zlevel":5}'
        '],'
        f'"series":{json.dumps(ec_series_mx, ensure_ascii=False)}'
        '}'
    )
    js_mx = js_mx.replace('"__SYMSIZE__"', 'function(val){return val[2];}')
    ec_html(js_mx, height=540)

with tab_pareto:
    gc_p = global_cc.copy()
    gc_p["Exceso_MM"] = ((gc_p["Monto_OC"]-gc_p["Monto_Ppto"]).clip(lower=0)/1e6)
    gc_p["Ahorro_MM"] = ((gc_p["Monto_Ppto"]-gc_p["Monto_OC"]).clip(lower=0)/1e6)
    exc = gc_p[gc_p["Exceso_MM"]>0].sort_values("Exceso_MM", ascending=False).copy()
    aho = gc_p[gc_p["Ahorro_MM"]>0].sort_values("Ahorro_MM", ascending=False).copy()
    tot_e = exc["Exceso_MM"].sum() if not exc.empty else 0.0
    tot_a = aho["Ahorro_MM"].sum() if not aho.empty else 0.0
    _n_rows  = max(len(exc), len(aho), 1)
    _chart_h = max(340, _n_rows * 50 + 90)
    _max_val = max(exc["Exceso_MM"].max() if not exc.empty else 0,
                   aho["Ahorro_MM"].max() if not aho.empty else 0, 0.1)

    _pc1, _pc2 = st.columns(2)

    def _ec_hbar_v2(labels, values, color_from, color_to, shadow_color, x_max, height):
        grad = {"type":"linear","x":0,"y":0,"x2":1,"y2":0,
                "colorStops":[{"offset":0,"color":color_from},{"offset":1,"color":color_to}]}
        opt = {
            "backgroundColor": "transparent",
            "animation": True, "animationDuration": 700, "animationEasing": "cubicOut",
            "tooltip": {
                "trigger": "axis",
                "backgroundColor": "rgba(10,15,30,0.97)",
                "borderColor": "rgba(255,255,255,0.12)", "borderWidth": 1,
                "textStyle": {"color": "#F1F5F9", "fontSize": 12},
                "extraCssText": "border-radius:10px;",
                "axisPointer": {"type": "shadow", "shadowStyle": {"color": "rgba(255,255,255,0.03)"}}
            },
            "toolbox": {"feature": {"saveAsImage": {"title": "Guardar"}},
                        "iconStyle": {"borderColor": "#334155"}, "right": 6, "top": 6},
            "grid": {"left": "2%", "right": "14%", "top": "3%", "bottom": "6%", "containLabel": True},
            "xAxis": {
                "type": "value", "max": round(float(x_max), 3),
                "axisLine": {"lineStyle": {"color": "rgba(255,255,255,0.08)"}},
                "splitLine": {"lineStyle": {"color": "rgba(255,255,255,0.04)", "type": "dashed"}},
                "axisLabel": {"color": "#64748B", "fontSize": 10, "formatter": "__FMT_X__"}
            },
            "yAxis": {
                "type": "category", "data": labels,
                "axisLine": {"lineStyle": {"color": "rgba(255,255,255,0.08)"}},
                "axisLabel": {"color": "#94A3B8", "fontSize": 11, "fontWeight": "600"},
                "splitLine": {"show": False}
            },
            "series": [{
                "type": "bar",
                "data": [round(float(v), 3) for v in values],
                "barMaxWidth": 28,
                "itemStyle": {
                    "color": grad,
                    "borderRadius": [0, 6, 6, 0],
                    "shadowBlur": 8,
                    "shadowColor": shadow_color,
                    "shadowOffsetX": 2,
                },
                "emphasis": {"itemStyle": {"shadowBlur": 18, "shadowColor": shadow_color}},
                "label": {
                    "show": True, "position": "right",
                    "color": "#94A3B8", "fontSize": 10, "fontWeight": "600",
                    "formatter": "__FMT_L__"
                },
            }]
        }
        js = json.dumps(opt, ensure_ascii=False)
        js = js.replace('"__FMT_X__"', '"MM$ {value}"')
        js = js.replace('"__FMT_L__"', 'function(p){return "MM$ "+p.value.toFixed(2);}')
        return js

    with _pc1:
        st.markdown(f"""
        <div style="background:linear-gradient(135deg,rgba(255,71,87,0.16),rgba(255,71,87,0.04));
                    border:1.5px solid rgba(255,71,87,0.45);border-radius:14px;
                    padding:14px 18px;margin-bottom:14px;
                    box-shadow:0 0 24px rgba(255,71,87,0.18);">
          <div style="font-size:0.60rem;font-weight:800;letter-spacing:0.18em;text-transform:uppercase;
                      color:rgba(255,71,87,0.85);margin-bottom:6px;">🔴 Cuentas con Sobrecoste</div>
          <div style="display:flex;align-items:baseline;gap:10px;">
            <div style="font-size:1.60rem;font-weight:900;color:#FF4757;text-shadow:0 0 16px rgba(255,71,87,0.5);">
              MM$ {_n(tot_e,2)}</div>
            <div style="font-size:0.72rem;color:#94A3B8;">{len(exc)} {"cuenta" if len(exc)==1 else "cuentas"} sobre ppto</div>
          </div>
        </div>""", unsafe_allow_html=True)
        if exc.empty:
            st.success("Ninguna cuenta supera su presupuesto.")
        else:
            exc_s = exc.sort_values("Exceso_MM", ascending=True)
            ec_html(_ec_hbar_v2(
                exc_s["CC_Label"].tolist(), exc_s["Exceso_MM"].tolist(),
                "rgba(255,71,87,0.95)", "rgba(255,120,100,0.55)",
                "rgba(255,71,87,0.40)", _max_val * 1.30, _chart_h
            ), height=_chart_h + 20)
            _pbl = exc[["CC_Label","Exceso_MM","Desv_Fin_%"]].copy()
            _pbl.columns = ["Cuenta de Costo","Sobrecoste (MM$)","Desvío (%)"]
            _maxe = _pbl["Sobrecoste (MM$)"].max() or 1
            def _ce(v): return f"background:linear-gradient(90deg,rgba(255,71,87,{v/_maxe*0.45:.2f}),transparent);color:#F8FAFC;font-weight:700"
            def _cde(v): return "color:#FF4757;font-weight:800"
            st.dataframe(
                _pbl.style
                    .format({"Sobrecoste (MM$)": _F2, "Desvío (%)": _FPp})
                    .map(_ce, subset=["Sobrecoste (MM$)"])
                    .map(_cde, subset=["Desvío (%)"]),
                use_container_width=True, hide_index=True)

    with _pc2:
        st.markdown(f"""
        <div style="background:linear-gradient(135deg,rgba(35,209,96,0.14),rgba(35,209,96,0.03));
                    border:1.5px solid rgba(35,209,96,0.40);border-radius:14px;
                    padding:14px 18px;margin-bottom:14px;
                    box-shadow:0 0 24px rgba(35,209,96,0.15);">
          <div style="font-size:0.60rem;font-weight:800;letter-spacing:0.18em;text-transform:uppercase;
                      color:rgba(35,209,96,0.85);margin-bottom:6px;">🟢 Cuentas con Ahorro</div>
          <div style="display:flex;align-items:baseline;gap:10px;">
            <div style="font-size:1.60rem;font-weight:900;color:#23D160;text-shadow:0 0 16px rgba(35,209,96,0.4);">
              MM$ {_n(tot_a,2)}</div>
            <div style="font-size:0.72rem;color:#94A3B8;">{len(aho)} {"cuenta" if len(aho)==1 else "cuentas"} bajo ppto</div>
          </div>
        </div>""", unsafe_allow_html=True)
        if aho.empty:
            st.warning("Ninguna cuenta está bajo presupuesto.")
        else:
            aho_s = aho.sort_values("Ahorro_MM", ascending=True)
            ec_html(_ec_hbar_v2(
                aho_s["CC_Label"].tolist(), aho_s["Ahorro_MM"].tolist(),
                "rgba(35,209,96,0.92)", "rgba(80,230,140,0.50)",
                "rgba(35,209,96,0.35)", _max_val * 1.30, _chart_h
            ), height=_chart_h + 20)
            _abl = aho[["CC_Label","Ahorro_MM","Desv_Fin_%"]].copy()
            _abl.columns = ["Cuenta de Costo","Ahorro (MM$)","Desvío (%)"]
            _maxa = _abl["Ahorro (MM$)"].max() or 1
            def _ca(v): return f"background:linear-gradient(90deg,rgba(35,209,96,{v/_maxa*0.38:.2f}),transparent);color:#F8FAFC;font-weight:700"
            def _cda(v): return "color:#23D160;font-weight:800"
            st.dataframe(
                _abl.style
                    .format({"Ahorro (MM$)": _F2, "Desvío (%)": _FPp})
                    .map(_ca, subset=["Ahorro (MM$)"])
                    .map(_cda, subset=["Desvío (%)"]),
                use_container_width=True, hide_index=True)

    st.markdown('<div style="height:16px;"></div>', unsafe_allow_html=True)
    _dif_neta    = tot_a - tot_e
    _dif_resumen = _saldo_fin / 1e6
    _match       = abs(_dif_neta - _dif_resumen) < 0.05
    _match_color = "#23D160" if _match else "#FFB300"
    _match_txt   = "✓ Cuadra con resumen" if _match else "⚠ Diferencia con resumen"
    _dif_color   = "#23D160" if _dif_neta >= 0 else "#FF4757"
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,rgba(84,112,198,0.10),rgba(84,112,198,0.03));
                border:1.5px solid rgba(84,112,198,0.28);border-radius:16px;
                padding:20px 24px;display:grid;grid-template-columns:repeat(4,1fr);gap:0;
                box-shadow:0 0 40px rgba(84,112,198,0.10),inset 0 1px 0 rgba(255,255,255,0.05);">
      <div style="text-align:center;border-right:1px solid rgba(255,255,255,0.07);padding-right:20px;">
        <div style="font-size:0.58rem;font-weight:800;letter-spacing:0.16em;text-transform:uppercase;color:rgba(255,71,87,0.70);margin-bottom:8px;">Total Sobrecoste</div>
        <div style="font-size:1.55rem;font-weight:900;color:#FF4757;text-shadow:0 0 14px rgba(255,71,87,0.45);">{_n(tot_e,2,"MM$ ")}</div>
        <div style="font-size:0.68rem;color:#475569;margin-top:4px;">{len(exc)} cuentas en rojo</div>
      </div>
      <div style="text-align:center;border-right:1px solid rgba(255,255,255,0.07);padding:0 20px;">
        <div style="font-size:0.58rem;font-weight:800;letter-spacing:0.16em;text-transform:uppercase;color:rgba(35,209,96,0.70);margin-bottom:8px;">Total Ahorro</div>
        <div style="font-size:1.55rem;font-weight:900;color:#23D160;text-shadow:0 0 14px rgba(35,209,96,0.40);">{_n(tot_a,2,"MM$ ")}</div>
        <div style="font-size:0.68rem;color:#475569;margin-top:4px;">{len(aho)} cuentas en verde</div>
      </div>
      <div style="text-align:center;border-right:1px solid rgba(255,255,255,0.07);padding:0 20px;">
        <div style="font-size:0.58rem;font-weight:800;letter-spacing:0.16em;text-transform:uppercase;color:#475569;margin-bottom:8px;">Diferencia Neta</div>
        <div style="font-size:1.55rem;font-weight:900;color:{_dif_color};text-shadow:0 0 14px {_dif_color}44;">{_n(_dif_neta,2,"MM$ ",sgn=True)}</div>
        <div style="font-size:0.68rem;color:#475569;margin-top:4px;">saldo calculado desde CC</div>
      </div>
      <div style="text-align:center;padding-left:20px;">
        <div style="font-size:0.58rem;font-weight:800;letter-spacing:0.16em;text-transform:uppercase;color:#475569;margin-bottom:8px;">Saldo Cuadro Resumen</div>
        <div style="font-size:1.55rem;font-weight:900;color:#F1F5F9;">{_n(_dif_resumen,2,"MM$ ",sgn=True)}</div>
        <div style="font-size:0.70rem;font-weight:800;color:{_match_color};margin-top:4px;">{_match_txt}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

st.divider()

# ═════════════════════════════════════════════════════════════════════════════
# AVANCE DE COMPRA
# ═════════════════════════════════════════════════════════════════════════════
seccion("📈","AVANCE DE COMPRA", color=C_PPTO)

def _build_avance_charts(df_grp, label_col):
    df_grp = df_grp[(df_grp["Monto_Ppto"]>0)|(df_grp["Monto_OC"]>0)].sort_values("Monto_Ppto",ascending=False).copy()
    if df_grp.empty:
        st.info("Sin datos para mostrar."); return
    labels  = df_grp[label_col].tolist()
    graf_h  = max(340, len(labels)*54 + 130)

    def _ec_avance_chart(vals_ppto, vals_oc, vals_rec, x_unit, title):
        p_d = [round(float(v), 2) for v in vals_ppto]
        o_d = [round(float(v), 2) for v in vals_oc]
        r_d = [round(float(v), 2) for v in vals_rec]
        _grad_ppto = {"type": "linear", "x": 0, "y": 0, "x2": 1, "y2": 0,
                      "colorStops": [{"offset": 0, "color": "rgba(120,100,255,0.88)"},
                                     {"offset": 1, "color": "rgba(84,112,198,0.32)"}]}
        _grad_oc   = {"type": "linear", "x": 0, "y": 0, "x2": 1, "y2": 0,
                      "colorStops": [{"offset": 0, "color": "rgba(255,100,60,1.0)"},
                                     {"offset": 1, "color": "rgba(255,185,0,0.95)"}]}
        _grad_rec  = {"type": "linear", "x": 0, "y": 0, "x2": 1, "y2": 0,
                      "colorStops": [{"offset": 0, "color": "rgba(35,209,96,1.0)"},
                                     {"offset": 1, "color": "rgba(35,209,96,0.58)"}]}
        opt = {
            "backgroundColor": "transparent",
            "animation": True, "animationDuration": 900, "animationEasing": "cubicOut",
            "tooltip": {"trigger": "axis", "backgroundColor": "rgba(6,11,21,0.97)",
                        "borderColor": "rgba(56,189,248,0.25)",
                        "borderRadius": 12,
                        "textStyle": {"color": "#F1F5F9", "fontSize": 11},
                        "axisPointer": {"type": "shadow",
                                        "shadowStyle": {"color": "rgba(56,189,248,0.05)"}}},
            "toolbox": {"feature": {"saveAsImage": {"title": "Guardar"},
                                    "restore": {"title": "Restaurar"}},
                        "iconStyle": {"borderColor": "#475569"}, "right": 5, "top": 5},
            "legend": {"data": ["Presupuestado", "OC Efectivo", "Recibido"],
                       "textStyle": {"color": "#94A3B8", "fontSize": 10},
                       "top": 5, "right": 80},
            "grid": {"left": 200, "right": 30, "top": 40, "bottom": 30},
            "xAxis": {"type": "value", "name": x_unit,
                      "nameTextStyle": {"color": "#64748B", "fontSize": 11},
                      "axisLine": {"lineStyle": {"color": "#1E293B"}},
                      "splitLine": {"lineStyle": {"color": "rgba(255,255,255,0.045)"}},
                      "axisLabel": {"color": "#94A3B8"}},
            "yAxis": {"type": "category", "data": labels, "inverse": True,
                      "axisLine": {"lineStyle": {"color": "#1E293B"}},
                      "axisLabel": {"color": "#94A3B8", "fontSize": 10},
                      "splitLine": {"show": False}},
            "series": [
                {"name": "Presupuestado", "type": "bar", "data": p_d, "barWidth": "55%",
                 "itemStyle": {"color": _grad_ppto, "borderRadius": [0, 4, 4, 0]},
                 "emphasis": {"itemStyle": {"shadowBlur": 14, "shadowColor": "rgba(120,100,255,0.55)"}},
                 "z": 1, "label": {"show": False}},
                {"name": "OC Efectivo", "type": "bar", "data": o_d, "barWidth": "38%",
                 "barGap": "-100%",
                 "itemStyle": {"color": _grad_oc, "borderRadius": [0, 4, 4, 0]},
                 "emphasis": {"itemStyle": {"shadowBlur": 14, "shadowColor": "rgba(255,100,60,0.60)"}},
                 "z": 2, "label": {"show": False}},
                {"name": "Recibido", "type": "bar", "data": r_d, "barWidth": "20%",
                 "barGap": "-100%",
                 "itemStyle": {"color": _grad_rec, "borderRadius": [0, 4, 4, 0]},
                 "emphasis": {"itemStyle": {"shadowBlur": 14, "shadowColor": "rgba(35,209,96,0.60)"}},
                 "z": 3, "label": {"show": False}},
            ]
        }
        return json.dumps(opt, ensure_ascii=False)

    _c1, _c2 = st.columns(2)
    with _c1:
        st.markdown('<div style="font-size:0.73rem;font-weight:700;letter-spacing:0.08em;color:#64748B;text-transform:uppercase;margin-bottom:6px;">💰 Financiero (MM$)</div>', unsafe_allow_html=True)
        ec_html(_ec_avance_chart(
            df_grp["Monto_Ppto"]/1e6, df_grp["Monto_OC"]/1e6, df_grp["Monto_Recibido"]/1e6, "MM$", "Financiero"
        ), height=graf_h)
    with _c2:
        st.markdown('<div style="font-size:0.73rem;font-weight:700;letter-spacing:0.08em;color:#64748B;text-transform:uppercase;margin-bottom:6px;">📦 Cantidades (m³)</div>', unsafe_allow_html=True)
        ec_html(_ec_avance_chart(
            df_grp["Cant_Ppto"], df_grp["Cant_OC"], df_grp["Cant_Recibida"], "m³", "Cantidades"
        ), height=graf_h)

tab_av_cc, tab_av_mat = st.tabs(["📂 Por Cuenta de Costo","🪨 Por Tipo de Árido"])
with tab_av_cc:
    _av_fc, _ = st.columns([1, 3])
    with _av_fc:
        _av_cc_opts = sorted(resumen["CC"].unique())
        _av_sel_cc  = filtro_checklist("Cuenta de Costo", _av_cc_opts, "av_cc", fmt=lambda x: cc_map.get(x, x))
    pcc = (resumen[resumen["CC"].isin(_av_sel_cc)]
           .groupby("CC").agg(Monto_Ppto=("Monto_Ppto","sum"), Monto_OC=("Monto_OC","sum"),
                               Monto_Recibido=("Monto_Recibido","sum"), Cant_Ppto=("Cant_Ppto","sum"),
                               Cant_OC=("Cant_OC","sum"), Cant_Recibida=("Cant_Recibida","sum")).reset_index())
    pcc["CC_Label"] = pcc["CC"].map(cc_map).fillna(pcc["CC"])
    _build_avance_charts(pcc, "CC_Label")

with tab_av_mat:
    _av_fm, _ = st.columns([1, 3])
    with _av_fm:
        _av_mat_opts = sorted(resumen["Material"].unique())
        _av_sel_mat  = filtro_checklist("Tipo de Árido", _av_mat_opts, "av_mat")
    pmt = (resumen[resumen["Material"].isin(_av_sel_mat)]
           .groupby("Material").agg(Monto_Ppto=("Monto_Ppto","sum"), Monto_OC=("Monto_OC","sum"),
                                     Monto_Recibido=("Monto_Recibido","sum"), Cant_Ppto=("Cant_Ppto","sum"),
                                     Cant_OC=("Cant_OC","sum"), Cant_Recibida=("Cant_Recibida","sum")).reset_index())
    _build_avance_charts(pmt, "Material")

st.divider()

# ═════════════════════════════════════════════════════════════════════════════
# EVOLUCIÓN TEMPORAL
# ═════════════════════════════════════════════════════════════════════════════
seccion("📆","EVOLUCIÓN TEMPORAL", color=C_RECIBIDO,
        badge_txt="Recepciones + Facturación aprobada", badge_color=C_RECIBIDO)

if df_recfac_f.empty:
    st.info("Sin datos de recepción y facturación para los filtros seleccionados.")
else:
    _APROBADOS_EV = {"Aprobada", "Totalmente Asociada"}

    _oc_monto = df_oc_f[["N_OC","CC","Material","Monto_Recibido"]].copy()
    _oc_monto["N_OC"] = _oc_monto["N_OC"].astype(str).str.strip()
    _oc_monto["_tot_oc"] = _oc_monto.groupby("N_OC")["Monto_Recibido"].transform("sum")
    _oc_monto["Prop_CC"] = (_oc_monto["Monto_Recibido"]
                            .div(_oc_monto["_tot_oc"].replace(0, float("nan"))).fillna(0))

    _rev = (df_recfac_f[df_recfac_f["N_Doc_Rec"] > 0]
            .drop_duplicates(["N_OC","N_Doc_Rec"])
            [["N_OC","N_Doc_Rec","Fecha_Rec","Monto_Rec"]].copy())

    # Misma fuente que RecFac KPI: Monto_Rec del archivo RecFac distribuido por CC
    df_rec_ev = (_rev[["N_OC","N_Doc_Rec","Fecha_Rec","Monto_Rec"]]
                 .merge(_oc_monto[["N_OC","CC","Material","Prop_CC"]], on="N_OC", how="left")
                 .dropna(subset=["Fecha_Rec"]).copy())
    df_rec_ev["Monto_Rec"] = (df_rec_ev["Monto_Rec"] * df_rec_ev["Prop_CC"]).fillna(0)
    df_rec_ev["CC_Label"]  = df_rec_ev["CC"].map(cc_map).fillna(df_rec_ev["CC"].fillna("Sin CC"))
    df_rec_ev["Material"]  = df_rec_ev["Material"].fillna("Sin Material")
    df_rec_ev["Mes"]       = df_rec_ev["Fecha_Rec"].dt.to_period("M").astype(str)

    _fac_base = (df_recfac_f[
                     (df_recfac_f["N_Factura"] > 0) &
                     (df_recfac_f["Estado_Doc_Fac"].astype(str).str.strip().isin(_APROBADOS_EV))
                 ].drop_duplicates("N_Factura")
                 [["N_OC","N_Factura","Monto_Factura","Fecha_Rec_Fac"]].copy())
    # Aplicar netting de notas de crédito antes de distribuir por CC
    _fac_base["Monto_Factura"] = (
        _fac_base["Monto_Factura"] - _fac_base["N_Factura"].map(nc_by_fac).fillna(0)
    ).clip(lower=0)
    df_fac_ev = (_fac_base
                 .merge(_oc_monto[["N_OC","CC","Material","Prop_CC"]], on="N_OC", how="left")
                 .dropna(subset=["Fecha_Rec_Fac"]).copy())
    df_fac_ev["Monto_Factura"] = (df_fac_ev["Monto_Factura"] * df_fac_ev["Prop_CC"]).fillna(0)
    df_fac_ev["CC_Label"]  = df_fac_ev["CC"].map(cc_map).fillna(df_fac_ev["CC"].fillna("Sin CC"))
    df_fac_ev["Material"]  = df_fac_ev["Material"].fillna("Sin Material")
    df_fac_ev["Mes"]       = df_fac_ev["Fecha_Rec_Fac"].dt.to_period("M").astype(str)

    _all_dates = pd.concat([df_rec_ev["Fecha_Rec"].rename("f"),
                             df_fac_ev["Fecha_Rec_Fac"].rename("f")]).dropna()
    if not _all_dates.empty:
        _fmin_ev = _all_dates.min().date(); _fmax_ev = _all_dates.max().date()
        _rango_ev = st.date_input("Rango de fechas:", value=(_fmin_ev, _fmax_ev),
                                   min_value=_fmin_ev, max_value=_fmax_ev, key="ev_fecha")
    else:
        _rango_ev = None

    if _rango_ev and len(_rango_ev) == 2:
        df_rec_ev = df_rec_ev[df_rec_ev["Fecha_Rec"].dt.date.between(_rango_ev[0], _rango_ev[1])]
        df_fac_ev = df_fac_ev[df_fac_ev["Fecha_Rec_Fac"].dt.date.between(_rango_ev[0], _rango_ev[1])]

    def _build_evolucion(df_rec, df_fac, group_col, paleta):
        rec_grp = df_rec.groupby(["Mes", group_col])["Monto_Rec"].sum().reset_index()
        fac_grp = df_fac.groupby(["Mes", group_col])["Monto_Factura"].sum().reset_index()
        all_months = sorted(set(rec_grp["Mes"].tolist() + fac_grp["Mes"].tolist()))
        grupos     = sorted(set(rec_grp[group_col].dropna().tolist() + fac_grp[group_col].dropna().tolist()))
        if not all_months:
            st.info("Sin datos en el período seleccionado."); return

        rec_tot = rec_grp.groupby("Mes")["Monto_Rec"].sum().reindex(all_months, fill_value=0)
        fac_tot = fac_grp.groupby("Mes")["Monto_Factura"].sum().reindex(all_months, fill_value=0)

        ec_series = []
        legend_data = []
        for i, grupo in enumerate(grupos):
            color = paleta[i % len(paleta)]
            sub_rec = (rec_grp[rec_grp[group_col] == grupo]
                       .set_index("Mes")["Monto_Rec"]
                       .reindex(all_months, fill_value=0))
            sub_fac = (fac_grp[fac_grp[group_col] == grupo]
                       .set_index("Mes")["Monto_Factura"]
                       .reindex(all_months, fill_value=0))
            legend_data.append(grupo)
            ec_series.append({
                "name": grupo, "type": "bar", "stack": "rec",
                "yAxisIndex": 0,
                "data": [round(float(v)/1e6, 3) for v in sub_rec.values],
                "itemStyle": {"color": color, "opacity": 0.88},
            })
            ec_series.append({
                "name": grupo + " · Fac.", "type": "bar", "stack": "fac",
                "yAxisIndex": 0,
                "data": [round(float(v)/1e6, 3) for v in sub_fac.values],
                "itemStyle": {"color": color, "opacity": 0.38},
            })

        _grad_rec_area = {"type": "linear", "x": 0, "y": 0, "x2": 0, "y2": 1,
                          "colorStops": [{"offset": 0, "color": "rgba(56,189,248,0.28)"},
                                         {"offset": 1, "color": "rgba(56,189,248,0.02)"}]}
        _grad_fac_area = {"type": "linear", "x": 0, "y": 0, "x2": 0, "y2": 1,
                          "colorStops": [{"offset": 0, "color": "rgba(167,139,250,0.22)"},
                                         {"offset": 1, "color": "rgba(167,139,250,0.02)"}]}
        ec_series.append({
            "name": "Acum. Recepcionado", "type": "line", "yAxisIndex": 1,
            "data": [round(float(v)/1e6, 3) for v in rec_tot.cumsum().values],
            "smooth": 0.5,
            "lineStyle": {"color": "#38BDF8", "width": 2.5,
                          "shadowBlur": 8, "shadowColor": "rgba(56,189,248,0.4)"},
            "itemStyle": {"color": "#38BDF8",
                          "shadowBlur": 6, "shadowColor": "rgba(56,189,248,0.5)"},
            "symbol": "circle", "symbolSize": 6,
            "areaStyle": {"color": _grad_rec_area}
        })
        ec_series.append({
            "name": "Acum. Facturación", "type": "line", "yAxisIndex": 1,
            "data": [round(float(v)/1e6, 3) for v in fac_tot.cumsum().values],
            "smooth": 0.5,
            "lineStyle": {"color": "#A78BFA", "width": 2.5, "type": "dashed",
                          "shadowBlur": 6, "shadowColor": "rgba(167,139,250,0.35)"},
            "itemStyle": {"color": "#A78BFA"},
            "symbol": "circle", "symbolSize": 6,
            "areaStyle": {"color": _grad_fac_area}
        })

        legend_all = legend_data + ["Acum. Recepcionado", "Acum. Facturación"]
        _h = max(460, len(all_months) * 22 + 220)

        opt_ev = {
            "backgroundColor": "transparent",
            "animation": True, "animationDuration": 900, "animationEasing": "cubicOut",
            "tooltip": {"trigger": "axis", "backgroundColor": "rgba(6,11,21,0.97)",
                        "borderColor": "rgba(56,189,248,0.22)",
                        "borderRadius": 12,
                        "textStyle": {"color": "#F1F5F9", "fontSize": 11},
                        "axisPointer": {"type": "shadow"}},
            "toolbox": {"feature": {"saveAsImage": {"title": "Guardar"},
                                    "restore": {"title": "Restaurar"},
                                    "dataView": {"readOnly": True, "title": "Ver datos"}},
                        "iconStyle": {"borderColor": "#475569"}, "right": 10, "top": 5},
            "legend": {"type": "scroll", "data": legend_all,
                       "textStyle": {"color": "#94A3B8", "fontSize": 10}, "bottom": 40},
            "dataZoom": [
                {"type": "inside"},
                {"type": "slider", "height": 20, "bottom": 10,
                 "fillerColor": "rgba(56,189,248,0.18)",
                 "handleStyle": {"color": "#38BDF8"}}
            ],
            "grid": {"top": 50, "bottom": 100, "left": 65, "right": 75},
            "xAxis": {"type": "category", "data": all_months,
                      "axisLine": {"lineStyle": {"color": "#1E293B"}},
                      "axisLabel": {"color": "#94A3B8", "rotate": -30},
                      "splitLine": {"lineStyle": {"color": "rgba(255,255,255,0.055)"}}},
            "yAxis": [
                {"type": "value", "name": "MM$ del mes",
                 "nameTextStyle": {"color": "#64748B", "fontSize": 11},
                 "axisLine": {"lineStyle": {"color": "#1E293B"}},
                 "splitLine": {"lineStyle": {"color": "rgba(255,255,255,0.055)"}},
                 "axisLabel": {"color": "#94A3B8"}},
                {"type": "value", "name": "MM$ acumulado",
                 "nameTextStyle": {"color": "#64748B", "fontSize": 11},
                 "axisLine": {"lineStyle": {"color": "#1E293B"}},
                 "splitLine": {"show": False},
                 "axisLabel": {"color": "#64748B"}}
            ],
            "series": ec_series
        }
        ec_html(json.dumps(opt_ev, ensure_ascii=False), height=_h)

        _tot_rec = df_rec["Monto_Rec"].sum()
        _tot_fac = df_fac["Monto_Factura"].sum()
        _n_rec   = df_rec["N_Doc_Rec"].nunique()
        _n_fac   = df_fac["N_Factura"].nunique()
        st.caption("↑ ▓ Barra sólida = Recepcionado · ░ Barra clara = Facturado Neto (−NC) · Líneas = Acumulado (eje der.)")
        _ev1, _ev2, _ev3, _ev4 = st.columns(4)
        _ev1.metric("Total Recepcionado",        _n(_tot_rec/1e6, 2, "MM$ "))
        _ev2.metric("Total Facturado Neto (−NC)", _n(_tot_fac/1e6, 2, "MM$ "))
        _ev3.metric("Docs de Recepción",          str(_n_rec))
        _ev4.metric("Facturas Aprobadas",          str(_n_fac))

    tab_ev_cc, tab_ev_mat = st.tabs(["📂 Por Cuenta de Costo","🪨 Por Tipo de Árido"])
    with tab_ev_cc:
        _filt_col, _ = st.columns([1, 3])
        with _filt_col:
            _cc_opts   = sorted(set(df_rec_ev["CC_Label"].dropna().tolist() + df_fac_ev["CC_Label"].dropna().tolist()))
            _sel_ev_cc = filtro_checklist("Cuenta de Costo", _cc_opts, "ev_cc")
        _build_evolucion(df_rec_ev[df_rec_ev["CC_Label"].isin(_sel_ev_cc)],
                         df_fac_ev[df_fac_ev["CC_Label"].isin(_sel_ev_cc)],
                         "CC_Label", PALETA_CC)
    with tab_ev_mat:
        _filt_col2, _ = st.columns([1, 3])
        with _filt_col2:
            _mat_opts    = sorted(set(df_rec_ev["Material"].dropna().tolist() + df_fac_ev["Material"].dropna().tolist()))
            _sel_ev_mat  = filtro_checklist("Tipo de Árido", _mat_opts, "ev_mat")
        _build_evolucion(df_rec_ev[df_rec_ev["Material"].isin(_sel_ev_mat)],
                         df_fac_ev[df_fac_ev["Material"].isin(_sel_ev_mat)],
                         "Material", PALETA_CC)

st.divider()

# ═════════════════════════════════════════════════════════════════════════════
# DETALLE OPERACIONAL
# ═════════════════════════════════════════════════════════════════════════════
seccion("🔍","DETALLE OPERACIONAL", color="#9a60b4")
tab_res, tab_oc, tab_recfac = st.tabs([
    "📋 Resumen CC + Material",
    "📄 Órdenes de Compra",
    "📦🧾 Recepción y Facturación"
])

with tab_res:
    _tr_fc, _tr_fm, _ = st.columns([1, 1, 2])
    with _tr_fc:
        _tr_sel_cc  = filtro_checklist("Cuenta de Costo", sorted(resumen["CC"].unique()), "tr_cc",
                                        fmt=lambda x: cc_map.get(x, x))
    with _tr_fm:
        _tr_sel_mat = filtro_checklist("Tipo de Árido", sorted(resumen["Material"].unique()), "tr_mat")
    r4 = resumen[resumen["CC"].isin(_tr_sel_cc) & resumen["Material"].isin(_tr_sel_mat)]
    td2 = r4[["Semaforo","CC_Label","Material","Cant_Ppto","Cant_OC","Cant_Recibida","Avance_%",
               "Monto_Ppto","Monto_OC","Monto_Recibido","Desv_Cant_%","Desv_Monto_%"]].copy()
    td2.columns = ["","Cuenta de Costo","Material","Cant Ppto (m³)","Cant OC (m³)","Cant Recibida (m³)",
                   "Avance (%)","Monto Ppto ($)","Monto OC ($)","Monto Recibido ($)","Desv Cant (%)","Desv Monto (%)"]
    st.dataframe(td2.style.format({"Cant Ppto (m³)":_F0,"Cant OC (m³)":_F0,"Cant Recibida (m³)":_F0,
                                    "Avance (%)":_FP,"Monto Ppto ($)":_FS0,"Monto OC ($)":_FS0,
                                    "Monto Recibido ($)":_FS0,"Desv Cant (%)":_FPp,"Desv Monto (%)":_FPp},na_rep="-")
                 .map(color_desv, subset=["Desv Cant (%)","Desv Monto (%)"]),
                 use_container_width=True, hide_index=True)
    st.markdown('<div style="height:12px;"></div>', unsafe_allow_html=True)
    _dsv     = ((r4["Monto_Ppto"].sum()-r4["Monto_OC"].sum())/r4["Monto_Ppto"].sum()*100) if r4["Monto_Ppto"].sum()>0 else 0
    _saldo_f = r4["Monto_Ppto"].sum() - r4["Monto_OC"].sum()
    _rs1,_rs2,_rs3,_rs4 = st.columns(4)
    with _rs1: st.markdown(mini_kpi("MONTO PRESUPUESTO", f"MM$ {r4['Monto_Ppto'].sum()/1e6:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","), f"{len(r4)} filas filtradas"), unsafe_allow_html=True)
    with _rs2: st.markdown(mini_kpi("MONTO OC EFECTIVO", f"MM$ {r4['Monto_OC'].sum()/1e6:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","), f"Desv. {_dsv:+.1f}% vs ppto".replace(".",","), color=C_OK if _dsv>0 else C_CRITICO), unsafe_allow_html=True)
    with _rs3: st.markdown(mini_kpi("MONTO RECIBIDO",    f"MM$ {r4['Monto_Recibido'].sum()/1e6:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","), "suma recepciones"), unsafe_allow_html=True)
    with _rs4: st.markdown(mini_kpi("SALDO PPTO − OC",   f"MM$ {_saldo_f/1e6:+,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","), "positivo = disponible", color=C_OK if _saldo_f>0 else C_CRITICO), unsafe_allow_html=True)
    _saldo_c = r4["Cant_Ppto"].sum() - r4["Cant_OC"].sum()
    _dsv_c   = (_saldo_c / r4["Cant_Ppto"].sum() * 100) if r4["Cant_Ppto"].sum() > 0 else 0
    _rq1,_rq2,_rq3,_rq4 = st.columns(4)
    with _rq1: st.markdown(mini_kpi("CANT. PRESUPUESTO", f"{r4['Cant_Ppto'].sum():,.0f} m³".replace(".","\x00").replace(",",".").replace("\x00",","), None), unsafe_allow_html=True)
    with _rq2: st.markdown(mini_kpi("CANT. OC EFECTIVO", f"{r4['Cant_OC'].sum():,.0f} m³".replace(".","\x00").replace(",",".").replace("\x00",","), f"Desv. {_dsv_c:+.1f}%".replace(".",","), color=C_OK if _dsv_c>0 else C_CRITICO), unsafe_allow_html=True)
    with _rq3: st.markdown(mini_kpi("CANT. RECIBIDA",    f"{r4['Cant_Recibida'].sum():,.0f} m³".replace(".","\x00").replace(",",".").replace("\x00",","), None), unsafe_allow_html=True)
    with _rq4: st.markdown(mini_kpi("SALDO PPTO − OC",   f"{_saldo_c:+,.0f} m³".replace(".","\x00").replace(",",".").replace("\x00",","), "positivo = disponible", color=C_OK if _saldo_c>0 else C_CRITICO), unsafe_allow_html=True)

with tab_oc:
    _or1,_or2,_or3 = st.columns(3)
    with _or1: _oc_cc   = filtro_checklist("Cuenta de Costo:", sorted(df_oc_f["CC"].dropna().unique()), "oc_cc", fmt=lambda x: cc_map.get(x,x))
    with _or2: _oc_mat  = filtro_checklist("Material:", sorted(df_oc_f["Material"].dropna().unique()), "oc_mat")
    with _or3: _oc_est  = filtro_checklist("Estado:", sorted(df_oc_f["Estado"].dropna().unique()), "oc_est")
    _or4,_or5,_or6 = st.columns(3)
    with _or4: _oc_prov = filtro_checklist("Proveedor:", sorted(df_oc_f["Proveedor"].dropna().unique()), "oc_prov")
    with _or5:
        _fo = df_oc_f["Fecha"].dropna()
        if not _fo.empty:
            _oc_rango = st.date_input("Fecha OC:", value=(_fo.min().date(),_fo.max().date()),
                                       min_value=_fo.min().date(), max_value=_fo.max().date(), key="oc_fecha")
        else: _oc_rango = None
    with _or6: _oc_nb = st.text_input("N° OC:", placeholder="Ej: 12345", key="oc_n_oc")

    dod = df_oc_f.copy()
    dod = dod[dod["CC"].isin(_oc_cc) & dod["Material"].isin(_oc_mat) & dod["Estado"].isin(_oc_est) & dod["Proveedor"].isin(_oc_prov)]
    if _oc_rango and len(_oc_rango)==2:
        dod = dod[dod["Fecha"].dt.date.between(_oc_rango[0], _oc_rango[1])]
    if _oc_nb.strip():
        dod = dod[dod["N_OC"].astype(str).str.contains(_oc_nb.strip(), case=False, na=False)]
    dod = dod[["N_OC","Fecha","CC","Proveedor","Material","Unidad","Cantidad","PU","SubTotal",
               "Cant_Recibida","Monto_Recibido","Cant_Efectiva","Monto_Efectivo","Estado"]].copy()
    dod["CC"] = dod["CC"].map(cc_map).fillna(dod["CC"])
    dod.columns = ["N° OC","Fecha","Cuenta de Costo","Proveedor","Material","Ud",
                   "Cant Ordenada","P. Unit.","Monto Ordenado","Cant Recibida","Monto Recibido",
                   "Cant Efectiva","Monto Efectivo","Estado"]

    _sum_oc_ef = dod["Monto Efectivo"].sum()
    _sum_rec   = dod["Monto Recibido"].sum()
    _saldo_rec = _sum_oc_ef - _sum_rec
    _ejec_pct  = (_sum_rec / _sum_oc_ef * 100) if _sum_oc_ef > 0 else 0
    _n_oc_det  = dod["N° OC"].nunique()

    if "oc_filtro" not in st.session_state: st.session_state["oc_filtro"] = "todo"
    _sel = st.session_state.get("oc_filtro", "todo")
    if _sel == "recibido":  dod_view = dod[dod["Monto Recibido"] > 0].copy()
    elif _sel == "saldo":   dod_view = dod[dod["Monto Efectivo"] > dod["Monto Recibido"]].copy()
    else:                   dod_view = dod

    _sk1,_sk2,_sk3,_sk4 = st.columns(4)
    with _sk1:
        if st.button(f"OC EFECTIVO\nMM$ {_sum_oc_ef/1e6:,.1f}\n{_n_oc_det} órdenes".replace(".","\x00").replace(",",".").replace("\x00",","),
                     key="ocf_todo", use_container_width=True, type="primary" if _sel=="todo" else "secondary"):
            st.session_state["oc_filtro"] = "todo"; st.rerun()
    with _sk2:
        if st.button(f"RECEPCIONADO\nMM$ {_sum_rec/1e6:,.1f}\n{_ejec_pct:.1f}% del OC efectivo".replace(".","\x00").replace(",",".").replace("\x00",","),
                     key="ocf_rec", use_container_width=True, type="primary" if _sel=="recibido" else "secondary"):
            st.session_state["oc_filtro"] = "recibido"; st.rerun()
    with _sk3:
        if st.button(f"SALDO X RECIBIR\nMM$ {_saldo_rec/1e6:,.1f}\npendiente de recepción".replace(".","\x00").replace(",",".").replace("\x00",","),
                     key="ocf_saldo", use_container_width=True, type="primary" if _sel=="saldo" else "secondary"):
            st.session_state["oc_filtro"] = "saldo"; st.rerun()
    with _sk4:
        st.markdown(mini_kpi("N° LÍNEAS", f"{len(dod_view):,}".replace(".","\x00").replace(",",".").replace("\x00",","), "filas en vista actual"), unsafe_allow_html=True)

    st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)
    st.dataframe(dod_view.style.format({"Cant Ordenada":_F1,"P. Unit.":_FS0,"Monto Ordenado":_FS0,
                                         "Cant Recibida":_F1,"Monto Recibido":_FS0,
                                         "Cant Efectiva":_F1,"Monto Efectivo":_FS0},na_rep="-")
                 .map(color_estado, subset=["Estado"]),
                 use_container_width=True, hide_index=True)
    st.markdown('<div style="height:12px;"></div>', unsafe_allow_html=True)
    _voc_ef  = dod_view["Monto Efectivo"].sum()
    _voc_rec = dod_view["Monto Recibido"].sum()
    _voc_sal = max(_voc_ef - _voc_rec, 0)
    _voc_pct = (_voc_rec / _voc_ef * 100) if _voc_ef > 0 else 0
    _od1,_od2,_od3,_od4,_od5 = st.columns(5)
    with _od1: st.markdown(mini_kpi("OC EFECTIVO",     f"MM$ {_voc_ef/1e6:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","),  f"{dod_view['N° OC'].nunique()} órdenes"), unsafe_allow_html=True)
    with _od2: st.markdown(mini_kpi("RECEPCIONADO",    f"MM$ {_voc_rec/1e6:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","), f"{_voc_pct:.1f}% del OC".replace(".",",")), unsafe_allow_html=True)
    with _od3: st.markdown(mini_kpi("SALDO X RECIBIR", f"MM$ {_voc_sal/1e6:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","), "pendiente", color=C_ALERTA if _voc_sal>0 else C_OK), unsafe_allow_html=True)
    with _od4: st.markdown(mini_kpi("CANT. EFECTIVA",  f"{dod_view['Cant Efectiva'].sum():,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","),  "unidades en vista"), unsafe_allow_html=True)
    with _od5: st.markdown(mini_kpi("N° LÍNEAS",       f"{len(dod_view):,}".replace(".","\x00").replace(",",".").replace("\x00",","),         "filas en vista"), unsafe_allow_html=True)

with tab_recfac:
    if df_recfac_f.empty:
        st.info("Sin datos para las OC de áridos seleccionadas.")
    else:
        _oc_unicas   = df_recfac_f.drop_duplicates("N_OC")
        _n_sin_rec   = (_oc_unicas["Estado_Rec_OC"] == "Sin Recepciones").sum()
        _df_fac_k    = df_recfac_f[df_recfac_f["N_Factura"] > 0]
        _df_rec_docs = df_recfac_f[df_recfac_f["N_Doc_Rec"] > 0].drop_duplicates(["N_OC","N_Doc_Rec"])
        _tot_rec_mm  = _df_rec_docs["Monto_Rec"].sum() / 1e6
        _tot_fac_mm  = _df_fac_k.drop_duplicates("N_Factura")["Monto_Factura"].sum() / 1e6
        _n_rec_sin_fac = (_df_rec_docs["N_Factura"] <= 0).sum()
        _APROBADOS = {"Aprobada", "Totalmente Asociada"}
        _df_con_fac = _df_rec_docs[_df_rec_docs["N_Factura"] > 0]
        _n_rec_fac_no_apr = (~_df_con_fac["Estado_Doc_Fac"].astype(str).str.strip().isin(_APROBADOS)).sum()

        # Alerta: comparación POR FACTURA — Monto_Neto vs suma de recepciones de esa factura
        # (evita falsos positivos cuando una factura cubre múltiples docs de recepción)
        _tot_nc_mm = nc_by_fac.sum() / 1e6 if not nc_by_fac.empty else 0
        _fac_base_a = (df_recfac_f[df_recfac_f["N_Factura"] > 0]
                       .groupby("N_Factura").agg(
                           N_OC=("N_OC", "first"),
                           Monto_Factura=("Monto_Factura", "first"),
                           Sum_Rec_Fac=("Monto_Rec", "sum")
                       ).reset_index())
        _fac_base_a["Monto_NC"]   = _fac_base_a["N_Factura"].map(nc_by_fac).fillna(0)
        _fac_base_a["Monto_Neto"] = (_fac_base_a["Monto_Factura"] - _fac_base_a["Monto_NC"]).clip(lower=0)
        _fac_base_a["Desv_Fac"]   = _fac_base_a["Monto_Neto"] - _fac_base_a["Sum_Rec_Fac"]
        _invoices_alert    = set(_fac_base_a.loc[_fac_base_a["Desv_Fac"] > 1000, "N_Factura"])
        _ocs_fac_mayor_rec = set(_fac_base_a.loc[_fac_base_a["Desv_Fac"] > 1000, "N_OC"])
        _n_fac_mayor_rec   = len(_invoices_alert)

        _fc1, _fc2, _ = st.columns([1, 1, 2])
        with _fc1:
            _est_opts    = sorted(df_recfac_f["Estado_Rec_OC"].dropna().unique())
            _sel_est_rec = filtro_checklist("Estado Recepción", _est_opts, "recfac_est")
        with _fc2:
            _fac_opts    = ["Con factura","Sin factura"]
            _sel_fac     = filtro_checklist("Estado Factura", _fac_opts, "recfac_fac")

        df_rf = df_recfac_f[df_recfac_f["Estado_Rec_OC"].isin(_sel_est_rec)].copy()
        if "Con factura"  not in _sel_fac: df_rf = df_rf[df_rf["N_Factura"] <= 0]
        if "Sin factura"  not in _sel_fac: df_rf = df_rf[df_rf["N_Factura"] >  0]

        _rf = df_rf[[
            "N_OC","Proveedor","Estado_Rec_OC",
            "N_Doc_Rec","Monto_Rec","Fecha_Rec","Estado_Guia","Saldo_x_Rec",
            "N_Factura","Monto_Factura","Fecha_Rec_Fac","Estado_Doc_Fac","Estado_Asoc_Fac"
        ]].copy()
        _rf.columns = [
            "N° OC","Proveedor","Estado Recepción OC",
            "N° Doc. Recepción","Monto Recibido ($)","Fecha Recepción","Estado Guía","Saldo x Recibir ($)",
            "N° Factura","Monto Factura ($)","Fecha Recep. Factura","Estado Factura","Estado Asociación"
        ]
        # Alerta y desviación a nivel factura (solo en primera aparición de c/ factura)
        _fac_desv_map  = _fac_base_a.set_index("N_Factura")["Desv_Fac"].to_dict()
        _rf["Alerta"]  = _rf["N° Factura"].isin(_invoices_alert).map({True: "⚠ FAC>REC", False: ""})
        _rf["Desv. ($)"] = _rf["N° Factura"].map(_fac_desv_map).fillna(float("nan"))
        # mostrar Desv. solo en primera fila de cada factura
        _rf.loc[_rf["N° Factura"].gt(0) & _rf["N° Factura"].duplicated(), "Desv. ($)"] = float("nan")
        _rf = _rf[["Alerta"] + [c for c in _rf.columns if c not in ("Alerta", "Desv. ($)")] + ["Desv. ($)"]]

        if "rf_filtro" not in st.session_state: st.session_state["rf_filtro"] = "todo"
        _rsel = st.session_state.get("rf_filtro", "todo")
        if _rsel == "sin_rec":         _rf_view = _rf[_rf["Estado Recepción OC"] == "Sin Recepciones"].copy()
        elif _rsel == "sin_fac":       _rf_view = _rf[(_rf["N° Doc. Recepción"] > 0) & (_rf["N° Factura"] <= 0)].copy()
        elif _rsel == "fac_mayor_rec": _rf_view = _rf[_rf["N° OC"].isin(_ocs_fac_mayor_rec)].copy()
        elif _rsel == "sin_fac_apr":
            _rf_view = _rf[(_rf["N° Doc. Recepción"] > 0) & (_rf["N° Factura"] > 0) &
                           (~_rf["Estado Factura"].astype(str).str.strip().isin(_APROBADOS))].copy()
        else: _rf_view = _rf.copy()

        _rk1,_rk2,_rk3 = st.columns(3)
        with _rk1:
            if st.button(f"TODAS\n{len(_rf):,} filas".replace(".","\x00").replace(",",".").replace("\x00",","),
                         key="rf_todo", use_container_width=True, type="primary" if _rsel=="todo" else "secondary"):
                st.session_state["rf_filtro"] = "todo"; st.rerun()
        with _rk2:
            if st.button(f"SIN RECEPCIÓN OC\n{_n_sin_rec} órdenes\npendientes",
                         key="rf_sin_rec", use_container_width=True, type="primary" if _rsel=="sin_rec" else "secondary"):
                st.session_state["rf_filtro"] = "sin_rec"; st.rerun()
        with _rk3:
            if st.button(f"FAC > REC\n{_n_fac_mayor_rec} OCs\nfactura excede recepción",
                         key="rf_fac_mayor", use_container_width=True,
                         type="primary" if _rsel=="fac_mayor_rec" else "secondary"):
                st.session_state["rf_filtro"] = "fac_mayor_rec"; st.rerun()

        _tot_fac_neto_mm = (_fac_base_a["Monto_Neto"].sum() / 1e6) if not _fac_base_a.empty else _tot_fac_mm
        # _fac_base_a ahora agrupa por N_Factura; _tot_fac_mm sigue siendo el bruto desde _df_fac_k
        st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
        _rk5,_rk6,_rk6b,_rk7,_rk8 = st.columns(5)
        with _rk5: st.markdown(mini_kpi("TOTAL RECIBIDO",    f"MM$ {_tot_rec_mm:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","), "suma docs recepción"), unsafe_allow_html=True)
        with _rk6: st.markdown(mini_kpi("FAC. BRUTA",        f"MM$ {_tot_fac_mm:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","), "facturas aprobadas"), unsafe_allow_html=True)
        with _rk6b: st.markdown(mini_kpi("FAC. NETA (−NC)",  f"MM$ {_tot_fac_neto_mm:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","), f"NC: MM$ {_tot_nc_mm:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","), color=C_OK), unsafe_allow_html=True)
        with _rk7:
            if st.button(f"RECEP. SIN FACTURA\n{_n_rec_sin_fac} docs\nsin factura asociada",
                         key="rf_sin_fac", use_container_width=True, type="primary" if _rsel=="sin_fac" else "secondary"):
                st.session_state["rf_filtro"] = "sin_fac"; st.rerun()
        with _rk8:
            if st.button(f"FAC. NO APROBADA\n{_n_rec_fac_no_apr} docs\nsin aprobación",
                         key="rf_sin_fac_apr", use_container_width=True, type="primary" if _rsel=="sin_fac_apr" else "secondary"):
                st.session_state["rf_filtro"] = "sin_fac_apr"; st.rerun()

        # Tabla de visualización: mostrar Monto_Factura solo en la primera fila
        # de cada factura (IConstruye repite el total de la factura en cada doc de recepción)
        _rf_display = _rf_view.copy()
        _dup_fac = _rf_display["N° Factura"].gt(0) & _rf_display["N° Factura"].duplicated()
        _rf_display.loc[_dup_fac, "Monto Factura ($)"]    = float("nan")
        _rf_display.loc[_dup_fac, "Fecha Recep. Factura"] = pd.NaT
        _rf_display.loc[_dup_fac, "Estado Factura"]       = None
        _rf_display.loc[_dup_fac, "Estado Asociación"]    = None

        st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)
        def _row_fac_mayor(row):
            bg = "background-color:rgba(238,102,102,0.13);" if row["Alerta"] == "⚠ FAC>REC" else ""
            return [bg] * len(row)
        def _color_desv_fac(val):
            if pd.isna(val) or val == 0: return ""
            return f"color:{C_CRITICO};font-weight:bold" if val > 0 else f"color:{C_OK};font-weight:bold"

        st.dataframe(
            _rf_display.style
               .apply(_row_fac_mayor, axis=1)
               .format({"N° Doc. Recepción":_FI, "N° Factura":_FI,
                        "Monto Recibido ($)":_FS0, "Saldo x Recibir ($)":_FS0,
                        "Monto Factura ($)":_FS0,
                        "Desv. ($)": lambda x: "-" if pd.isna(x) else (f"+${x:,.0f}" if x > 0 else f"${x:,.0f}").replace(".","\x00").replace(",",".").replace("\x00",",")},
                       na_rep="-")
               .map(color_rec_oc,    subset=["Estado Recepción OC"])
               .map(color_fac,       subset=["Estado Guía","Estado Factura","Estado Asociación"])
               .map(color_alerta_fac, subset=["Alerta"])
               .map(_color_desv_fac,  subset=["Desv. ($)"]),
            use_container_width=True, hide_index=True
        )
        st.caption(f"{len(_rf_view):,} documentos · {_rf_view['N° OC'].nunique():,} OCs de áridos".replace(".","\x00").replace(",",".").replace("\x00",","))
        st.markdown('<div style="height:12px;"></div>', unsafe_allow_html=True)
        _rf_rec_v    = _rf_view[_rf_view["N° Doc. Recepción"] > 0].drop_duplicates("N° Doc. Recepción")
        _rf_fac_v    = _rf_view[_rf_view["N° Factura"] > 0].drop_duplicates("N° Factura")
        _rf_sin_fapr = _rf_rec_v[(_rf_rec_v["N° Factura"] <= 0) |
                                   (~_rf_rec_v["Estado Factura"].astype(str).str.strip().isin(_APROBADOS))]
        _monto_sin_fapr = _rf_sin_fapr["Monto Recibido ($)"].sum()
        _n_sin_fapr     = len(_rf_sin_fapr)
        _rf1,_rf2,_rf3,_rf4,_rf5 = st.columns(5)
        with _rf1: st.markdown(mini_kpi("MONTO RECIBIDO",          f"MM$ {_rf_rec_v['Monto Recibido ($)'].sum()/1e6:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","), f"{len(_rf_rec_v)} docs"), unsafe_allow_html=True)
        with _rf2: st.markdown(mini_kpi("MONTO FACTURADO",         f"MM$ {_rf_fac_v['Monto Factura ($)'].sum()/1e6:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","),  f"{len(_rf_fac_v)} facturas"), unsafe_allow_html=True)
        with _rf3: st.markdown(mini_kpi("SIN FAC. APROBADA (MM$)", f"MM$ {_monto_sin_fapr/1e6:,.1f}".replace(".","\x00").replace(",",".").replace("\x00",","), "sin factura aprobada", color=C_CRITICO if _monto_sin_fapr>0 else C_OK), unsafe_allow_html=True)
        with _rf4: st.markdown(mini_kpi("DOCS SIN FAC. APR.",      str(_n_sin_fapr), "recepciones pendientes", color=C_CRITICO if _n_sin_fapr>0 else C_OK), unsafe_allow_html=True)
        with _rf5: st.markdown(mini_kpi("FILAS EN VISTA",          f"{len(_rf_view):,}".replace(".","\x00").replace(",",".").replace("\x00",","), "documentos en vista"), unsafe_allow_html=True)

# ── AUTO-REFRESCO ─────────────────────────────────────────────────────────────
st_autorefresh(interval=INTERVALO * 1000, key="data_refresh")
