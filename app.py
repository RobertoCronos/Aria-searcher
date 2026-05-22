import streamlit as st
import google.generativeai as genai
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
import json, re, io
from datetime import datetime

st.set_page_config(page_title="ARIA Licencias", page_icon="🎓", layout="centered", initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Mono:wght@400;500&family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 0 !important; padding-bottom: 3rem; max-width: 860px; }

/* ── Background ── */
body { background: #EEF4FF; }
.stApp { background: linear-gradient(160deg, #1A7FE8 0%, #1565C8 30%, #EEF4FF 60%); min-height: 100vh; }

/* ── Hero header ── */
.aria-hero {
    padding: 52px 36px 44px;
    text-align: center;
    margin-bottom: 0;
}
.aria-eyebrow {
    display: inline-flex; align-items: center; gap: 7px;
    background: rgba(255,255,255,0.18);
    border: 1px solid rgba(255,255,255,0.35);
    border-radius: 100px; padding: 5px 16px;
    font-size: 11px; letter-spacing: 2px; color: rgba(255,255,255,0.9);
    text-transform: uppercase; font-weight: 600; margin-bottom: 18px;
}
.aria-title {
    font-family: 'DM Serif Display', serif;
    font-size: clamp(2.4rem, 6vw, 3.6rem);
    color: #ffffff;
    margin: 0 0 10px;
    line-height: 1.1;
    letter-spacing: -1px;
}
.aria-title em { font-style: italic; color: #BEE3FF; }
.aria-sub {
    font-size: 1rem; color: rgba(255,255,255,0.72);
    font-weight: 300; max-width: 480px; margin: 0 auto;
    line-height: 1.6;
}

/* ── Main panel ── */
.main-panel {
    background: #ffffff;
    border-radius: 24px;
    padding: 36px 40px;
    margin-top: 0;
    box-shadow: 0 20px 60px rgba(21,101,200,0.18), 0 4px 16px rgba(0,0,0,0.06);
}

/* ── Section labels ── */
.slabel {
    font-size: 10px; letter-spacing: 2.5px; text-transform: uppercase;
    color: #94A3B8; font-weight: 600; margin-bottom: 8px; display: block;
}

/* ── Drop zone ── */
.dropzone {
    border: 2px dashed #CBD5E1;
    border-radius: 14px;
    background: #F8FAFF;
    padding: 24px 20px;
    text-align: center;
    transition: border-color .2s;
}
.dropzone:hover { border-color: #1A7FE8; }
.dropzone-icon { font-size: 24px; margin-bottom: 6px; }
.dropzone-main { font-size: 14px; color: #64748B; }
.dropzone-main strong { color: #1A7FE8; }
.dropzone-hint { font-size: 11px; color: #94A3B8; margin-top: 4px; }

/* ── Stats bar ── */
.stats-bar {
    display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 24px;
}
.stat-pill {
    background: #EEF4FF;
    border: 1px solid #C7DEFF;
    border-radius: 100px;
    padding: 6px 16px;
    font-size: 13px; color: #1565C8; font-weight: 500;
    display: flex; align-items: center; gap: 6px;
}
.stat-pill span { font-weight: 700; color: #1A7FE8; }

/* ── Layer legend ── */
.layer-legend { display: flex; gap: 16px; margin-bottom: 20px; flex-wrap: wrap; }
.legend-item { display: flex; align-items: center; gap: 7px; font-size: 12px; color: #64748B; font-weight: 500; }
.legend-dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }

/* ── Software cards ── */
.sw-card {
    background: #FFFFFF;
    border: 1.5px solid #E2E8F0;
    border-radius: 16px;
    padding: 20px 22px;
    margin-bottom: 12px;
    transition: box-shadow .2s, border-color .2s;
}
.sw-card:hover { box-shadow: 0 4px 20px rgba(26,127,232,0.1); border-color: #A8C8FF; }
.sw-card.l1 { border-left: 4px solid #1A7FE8; }
.sw-card.l2 { border-left: 4px solid #10B981; }
.sw-card.l3 { border-left: 4px solid #F59E0B; }

.sw-name { font-size: 15px; font-weight: 600; color: #0F172A; margin-bottom: 4px; }
.sw-desc { font-size: 13px; color: #64748B; line-height: 1.5; margin-bottom: 8px; }
.sw-detail { font-size: 12px; color: #94A3B8; margin-bottom: 3px; }
.sw-detail strong { color: #475569; }

/* ── Layer badges ── */
.layer-badge {
    display: inline-block; border-radius: 100px; padding: 3px 11px;
    font-size: 10px; font-weight: 600; letter-spacing: 1px;
    text-transform: uppercase; margin-bottom: 8px; margin-right: 6px;
}
.layer-1 { background: #EEF4FF; color: #1A7FE8; border: 1px solid #C7DEFF; }
.layer-2 { background: #ECFDF5; color: #059669; border: 1px solid #A7F3D0; }
.layer-3 { background: #FFFBEB; color: #D97706; border: 1px solid #FDE68A; }

/* ── Score pill ── */
.score-pill-sm {
    display: inline-flex; align-items: center; gap: 4px;
    background: #F8FAFF; border: 1px solid #C7DEFF;
    border-radius: 100px; padding: 2px 10px;
    font-size: 11px; color: #1565C8; font-weight: 600;
}
.price-pill {
    display: inline-flex; align-items: center; gap: 4px;
    border-radius: 100px; padding: 2px 9px;
    font-size: 10px; font-weight: 600; margin-left: 2px;
}
.price-free { background: #ECFDF5; color: #059669; border: 0.5px solid #A7F3D0; }
.price-paid { background: #FFFBEB; color: #D97706; border: 0.5px solid #FDE68A; }
.price-inst { background: #F1F5F9; color: #64748B; border: 0.5px solid #CBD5E1; }
.score-row { display: flex; align-items: center; gap: 8px; margin: 8px 0 2px; }
.score-label { font-size: 10px; color: #94A3B8; font-weight: 500; white-space: nowrap; min-width: 52px; }
.score-track { flex: 1; height: 5px; background: #E2E8F0; border-radius: 3px; overflow: hidden; }
.score-fill { height: 100%; border-radius: 3px; }
.score-num { font-size: 11px; font-weight: 600; color: #334155; min-width: 28px; text-align: right; }
.card-divider { height: 0.5px; background: #E2E8F0; margin: 10px 0; }

/* ── Package cards ── */
.pkg-card {
    background: linear-gradient(135deg, #1565C8 0%, #1A7FE8 100%);
    border-radius: 18px; padding: 22px 24px; margin-bottom: 14px;
    box-shadow: 0 8px 24px rgba(26,127,232,0.25);
}
.pkg-title { font-size: 15px; font-weight: 700; color: #ffffff; margin-bottom: 5px; }
.pkg-obj { font-size: 12px; color: rgba(255,255,255,0.7); margin-bottom: 10px; }
.pkg-orgs { font-size: 12px; color: rgba(255,255,255,0.9); line-height: 1.6; }
.pkg-benefit { font-size: 11px; color: #BEE3FF; margin-top: 8px; font-weight: 500; }

/* ── Dork cards ── */
.dork-card {
    background: #F8FAFF;
    border: 1.5px solid #E2EEFF;
    border-radius: 12px; padding: 10px 16px;
    margin-bottom: 8px; display: flex; align-items: flex-start; gap: 12px;
}
.dork-num { font-size: 11px; color: #94A3B8; font-weight: 700; min-width: 20px; padding-top: 1px; }
.dork-text { font-family: 'DM Mono', monospace; font-size: 12px; color: #1565C8; line-height: 1.5; flex: 1; word-break: break-all; }

/* ── Divider ── */
.aria-divider { height: 1px; background: #E2E8F0; margin: 28px 0; }

/* ── Streamlit overrides ── */
.stButton > button {
    background: linear-gradient(135deg, #1565C8, #1A7FE8) !important;
    color: white !important; border: none !important;
    border-radius: 12px !important; padding: 0.7rem 2rem !important;
    font-family: 'Inter', sans-serif !important; font-weight: 600 !important;
    font-size: 15px !important; width: 100% !important;
    box-shadow: 0 4px 16px rgba(26,127,232,0.35) !important;
    transition: all .2s !important;
}
.stButton > button:hover { opacity: 0.92 !important; transform: translateY(-1px) !important; }

.stTextArea textarea {
    border-radius: 12px !important; border: 1.5px solid #E2E8F0 !important;
    font-family: 'Inter', sans-serif !important; font-size: 14px !important;
    background: #F8FAFF !important;
}
.stTextArea textarea:focus { border-color: #1A7FE8 !important; box-shadow: 0 0 0 3px rgba(26,127,232,0.12) !important; }

.stFileUploader {
    border: 2px dashed #CBD5E1 !important; border-radius: 14px !important;
    background: #F8FAFF !important; padding: 8px !important;
}

.stDownloadButton > button {
    background: linear-gradient(135deg, #059669, #10B981) !important;
    color: white !important; border: none !important;
    border-radius: 12px !important; font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important; width: 100% !important;
    box-shadow: 0 4px 16px rgba(16,185,129,0.3) !important;
}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: #F1F5FE !important; border-radius: 12px !important;
    padding: 4px !important; gap: 2px !important;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 9px !important; font-family: 'Inter', sans-serif !important;
    font-weight: 500 !important; font-size: 13px !important;
    color: #64748B !important;
}
.stTabs [aria-selected="true"] {
    background: white !important; color: #1565C8 !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.1) !important;
}

/* ── Expander ── */
.streamlit-expanderHeader {
    background: #F8FAFF !important; border-radius: 10px !important;
    border: 1px solid #E2E8F0 !important; font-weight: 500 !important;
}

.stTextInput input {
    border-radius: 12px !important; border: 1.5px solid #E2E8F0 !important;
    background: #F8FAFF !important;
}

footer-note { text-align: center; font-size: 11px; color: #94A3B8; margin-top: 32px; }
</style>
""", unsafe_allow_html=True)


# ── Helpers ──────────────────────────────────────────────────────────
def extract_text(f):
    name = f.name.lower()
    if name.endswith('.txt'):
        return f.read().decode('utf-8', errors='ignore')
    elif name.endswith('.pdf'):
        try:
            import pypdf
            return '\n'.join(p.extract_text() or '' for p in pypdf.PdfReader(f).pages)
        except: return ""
    elif name.endswith('.docx'):
        try:
            import docx
            return '\n'.join(p.text for p in docx.Document(io.BytesIO(f.read())).paragraphs)
        except: return ""
    return ""


def parse_json(raw):
    try:
        clean = re.sub(r'```[a-z]*', '', raw).replace('`','').strip()
        start = clean.index('{')
        end   = clean.rindex('}') + 1
        return json.loads(clean[start:end])
    except:
        return None


SYSTEM_PROMPT = """Actua como consultor senior McKinsey en portafolios tecnologicos universitarios de ciencias de la salud.

MODELO DE TRES CAPAS:
CAPA 1: Herramientas core directamente alineadas a la disciplina central del dossier.
CAPA 2: Herramientas de disciplinas relacionadas (calidad, estadistica, RRHH, acreditacion).
CAPA 3: Ecosistema estrategico (gestion hospitalaria, simulacion, liderazgo, investigacion).

FILTROS DE INCLUSION (todos obligatorios):
- Licencia academica o institucional (no acceso publico general)
- Acceso restringido por afiliacion universitaria o convenio formal
- Uso real verificable en academia o industria de salud
- Reconocimiento internacional verificable

FILTROS DE EXCLUSION:
- Software gratuito sin control de acceso
- Trials menores a 12 meses
- Plataformas freemium sin restriccion academica real

CONDICIONES: Minimo 50 herramientas, sin duplicados, solo URLs verificables, juicio critico.

Responde SOLO con JSON valido, SIN bloques de codigo markdown, SIN texto antes o despues:
{"portafolio":[{"nombre":"string","capa":1,"tipo_acceso":"string","url_oficial":"string","relacion_dossier":"string","valor_estrategico":"string","certificacion":"string","paquete":"Hospitales|Medicina y Clinica|Cuidados Medicos|Cuidados Intensivos|Investigacion y Publicacion|Direccion y Liderazgo|Salud Publica y Politicas","score_academico":4,"score_cv":4,"score_industria":4,"score_diferenciacion":4}],"conceptos":[{"palabra_clave":"string","tecnologias_asociadas":"string","metodos_marcos":"string","areas_aplicacion":"string"}],"google_dorks":["dork1"]}"""


def call_gemini(api_key, dossier, context):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.5-flash", system_instruction=SYSTEM_PROMPT)
    msg = f"DOSSIER:\n{dossier[:6000] if dossier else 'Sin dossier.'}\n\nCONTEXTO:\n{context or 'Posgrado en ciencias de la salud.'}\n\nResponde SOLO con el JSON sin markdown."
    raw = model.generate_content(msg).text
    result = parse_json(raw)
    return result if result else {"portafolio":[], "conceptos":[], "google_dorks":[], "_raw": raw[:800]}


# ── Excel ─────────────────────────────────────────────────────────────
CAPA_COLORS = {1:"1A7FE8", 2:"059669", 3:"D97706"}
CAPA_BG     = {1:"EEF4FF", 2:"ECFDF5", 3:"FFFBEB"}
PAQUETES    = ["Hospitales","Medicina y Clinica","Cuidados Medicos","Cuidados Intensivos",
               "Investigacion y Publicacion","Direccion y Liderazgo","Salud Publica y Politicas"]

def _h(ws,r,c,v,bg="1565C8",fg="FFFFFF",sz=11):
    x=ws.cell(r,c,v); x.font=Font(bold=True,color=fg,name="Arial",size=sz)
    x.fill=PatternFill("solid",fgColor=bg); thin=Side(style="thin",color="CCCCCC")
    x.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    x.border=Border(left=thin,right=thin,top=thin,bottom=thin); return x

def _c(ws,r,c,v,bold=False,color="334155",bg=None,ha="left"):
    x=ws.cell(r,c,v); x.font=Font(bold=bold,color=color,name="Arial",size=10)
    x.alignment=Alignment(horizontal=ha,vertical="center",wrap_text=True)
    thin=Side(style="thin",color="E2E8F0"); x.border=Border(left=thin,right=thin,top=thin,bottom=thin)
    if bg: x.fill=PatternFill("solid",fgColor=bg)
    return x

def build_excel(data):
    wb=openpyxl.Workbook(); port=data.get("portafolio",[]); conc=data.get("conceptos",[]); dorks=data.get("google_dorks",[])

    ws1=wb.active; ws1.title="Portafolio"
    ws1.merge_cells("A1:I1"); tc=ws1.cell(1,1,"PORTAFOLIO ESTRATEGICO — ARIA Licencias")
    tc.font=Font(bold=True,color="FFFFFF",name="Arial",size=13); tc.fill=PatternFill("solid",fgColor="1565C8")
    tc.alignment=Alignment(horizontal="center",vertical="center"); ws1.row_dimensions[1].height=32
    for i,(h,w) in enumerate(zip(["#","Nombre","Tipo","Capa","Acceso","Link","Relacion Dossier","Valor Estrategico","Certificacion"],[4,24,18,8,24,32,36,36,30]),1):
        _h(ws1,2,i,h); ws1.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    ws1.row_dimensions[2].height=30
    for idx,sw in enumerate(port,1):
        row=idx+2; capa=sw.get("capa",1); alt="F8FAFF" if idx%2==0 else "FFFFFF"
        for col,val in enumerate([idx,sw.get("nombre",""),"Premium Institucional",f"Capa {capa}",sw.get("tipo_acceso",""),sw.get("url_oficial",""),sw.get("relacion_dossier",""),sw.get("valor_estrategico",""),sw.get("certificacion","")],1):
            _c(ws1,row,col,val,bold=(col==4),color=CAPA_COLORS.get(capa,"334155") if col==4 else ("1A7FE8" if col==6 else "334155"),bg=CAPA_BG.get(capa,alt) if col==4 else alt,ha="center" if col in(1,3,4) else "left")
        ws1.row_dimensions[row].height=52
    ws1.freeze_panes="A3"; ws1.auto_filter.ref=f"A2:I{len(port)+2}"

    ws2=wb.create_sheet("Paquetes"); ws2.merge_cells("A1:D1")
    tc2=ws2.cell(1,1,"PAQUETES ESTRATEGICOS"); tc2.font=Font(bold=True,color="FFFFFF",name="Arial",size=13)
    tc2.fill=PatternFill("solid",fgColor="1565C8"); tc2.alignment=Alignment(horizontal="center",vertical="center"); ws2.row_dimensions[1].height=32
    for i,(h,w) in enumerate(zip(["Paquete","Objetivo","Software","Beneficio"],[24,40,50,40]),1):
        _h(ws2,2,i,h); ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    paquetes={}
    for sw in port: paquetes.setdefault(sw.get("paquete","Otros"),[]).append(sw)
    pkg_bgs=["1565C8","1A7FE8","0369A1","059669","7C3AED","9A3412","065F46"]; row2=3
    for pi,pname in enumerate(PAQUETES):
        sws=paquetes.get(pname,[])
        if not sws: continue
        bg_col=pkg_bgs[pi%len(pkg_bgs)]; names=", ".join(s.get("nombre","") for s in sws[:12])
        for col,val in enumerate([pname,f"Herramientas de {pname} nivel institucional.",names,f"Fortalece CV en {pname}."],1):
            c=ws2.cell(row2,col,val); c.font=Font(bold=(col==1),color="FFFFFF" if col==1 else "0D1B2A",name="Arial",size=10)
            c.fill=PatternFill("solid",fgColor=bg_col if col==1 else ("F8FAFF" if row2%2==0 else "FFFFFF"))
            c.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True)
            thin=Side(style="thin",color="E2E8F0"); c.border=Border(left=thin,right=thin,top=thin,bottom=thin)
        ws2.row_dimensions[row2].height=60; row2+=1
    ws2.freeze_panes="A3"

    ws3=wb.create_sheet("Conceptos"); ws3.merge_cells("A1:D1")
    tc3=ws3.cell(1,1,"CONCEPTOS CLAVE — ARIA Licencias"); tc3.font=Font(bold=True,color="FFFFFF",name="Arial",size=13)
    tc3.fill=PatternFill("solid",fgColor="1565C8"); tc3.alignment=Alignment(horizontal="center",vertical="center"); ws3.row_dimensions[1].height=32
    for i,(h,w) in enumerate(zip(["Palabra Clave","Tecnologias","Metodos","Areas"],[36,40,36,36]),1):
        _h(ws3,2,i,h); ws3.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    for idx,c in enumerate(conc,1):
        row3=idx+2; bg3="EEF4FF" if idx%2==0 else "FFFFFF"
        for col,key in enumerate(["palabra_clave","tecnologias_asociadas","metodos_marcos","areas_aplicacion"],1):
            _c(ws3,row3,col,c.get(key,""),bg=bg3)
        ws3.row_dimensions[row3].height=40
    ws3.freeze_panes="A3"

    ws4=wb.create_sheet("Scoring"); ws4.merge_cells("A1:G1")
    tc4=ws4.cell(1,1,"SCORING ESTRATEGICO"); tc4.font=Font(bold=True,color="FFFFFF",name="Arial",size=13)
    tc4.fill=PatternFill("solid",fgColor="1565C8"); tc4.alignment=Alignment(horizontal="center",vertical="center"); ws4.row_dimensions[1].height=32
    for i,(h,w) in enumerate(zip(["#","Nombre","Relevancia","Valor CV","Industria","Diferenciacion","SCORE TOTAL"],[4,30,18,14,16,20,14]),1):
        _h(ws4,2,i,h); ws4.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    ws4.row_dimensions[2].height=36
    for idx,sw in enumerate(port,1):
        row4=idx+2; bg4="F8FAFF" if idx%2==0 else "FFFFFF"; capa=sw.get("capa",1)
        _c(ws4,row4,1,idx,ha="center",bg=bg4); _c(ws4,row4,2,sw.get("nombre",""),bold=True,color=CAPA_COLORS.get(capa,"334155"),bg=CAPA_BG.get(capa,bg4))
        for col,key in zip([3,4,5,6],["score_academico","score_cv","score_industria","score_diferenciacion"]):
            _c(ws4,row4,col,sw.get(key,3),ha="center",bg=bg4)
        tc=ws4.cell(row4,7,f"=SUM(C{row4}:F{row4})"); tc.font=Font(bold=True,name="Arial",size=10)
        tc.alignment=Alignment(horizontal="center",vertical="center"); thin=Side(style="thin",color="E2E8F0")
        tc.border=Border(left=thin,right=thin,top=thin,bottom=thin); tc.fill=PatternFill("solid",fgColor=bg4)
        ws4.row_dimensions[row4].height=22
    if port:
        ws4.conditional_formatting.add(f"G3:G{len(port)+2}",ColorScaleRule(start_type="min",start_color="FF4444",mid_type="percentile",mid_value=50,mid_color="FFBB33",end_type="max",end_color="00C851"))
        for cl in ["C","D","E","F"]:
            ws4.conditional_formatting.add(f"{cl}3:{cl}{len(port)+2}",ColorScaleRule(start_type="min",start_color="FFCCCC",end_type="max",end_color="CCFFCC"))
    ws4.freeze_panes="A3"

    ws5=wb.create_sheet("Google Dorks"); ws5.merge_cells("A1:C1")
    tc5=ws5.cell(1,1,"GOOGLE DORKS — ARIA Licencias"); tc5.font=Font(bold=True,color="FFFFFF",name="Arial",size=13)
    tc5.fill=PatternFill("solid",fgColor="059669"); tc5.alignment=Alignment(horizontal="center",vertical="center"); ws5.row_dimensions[1].height=32
    for i,(h,w) in enumerate(zip(["#","Dork","Uso"],[6,80,40]),1):
        _h(ws5,2,i,h,bg="059669"); ws5.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    for idx,dork in enumerate(dorks,1):
        row5=idx+2; bg5="ECFDF5" if idx%2==0 else "FFFFFF"
        _c(ws5,row5,1,idx,ha="center",bg=bg5)
        c=ws5.cell(row5,2,dork); c.font=Font(name="Courier New",size=10,color="059669")
        c.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True)
        thin=Side(style="thin",color="E2E8F0"); c.border=Border(left=thin,right=thin,top=thin,bottom=thin); c.fill=PatternFill("solid",fgColor=bg5)
        _c(ws5,row5,3,"",bg=bg5); ws5.row_dimensions[row5].height=22
    ws5.freeze_panes="A3"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf.getvalue()


# ── UI ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="aria-hero">
  <div class="aria-eyebrow">✦ Herramienta Académica</div>
  <h1 class="aria-title">ARIA <em>Licencias</em></h1>
  <p class="aria-sub">Portafolio estratégico de software con licencia académica para programas de posgrado</p>
</div>
""", unsafe_allow_html=True)

# ── Contenedor blanco principal
with st.container():
    st.markdown('<div class="main-panel">', unsafe_allow_html=True)

    # API Key
    if "gemini_key" not in st.session_state:
        secret_key = st.secrets.get("GEMINI_API_KEY", "")
        if secret_key:
            st.session_state["gemini_key"] = secret_key

    if not st.session_state.get("gemini_key"):
        with st.expander("🔑 Configurar API Key de Gemini", expanded=True):
            k = st.text_input("Google Gemini API Key", type="password", placeholder="AIza...")
            if k:
                st.session_state["gemini_key"] = k
                st.success("✅ API Key guardada")

    api_key = st.session_state.get("gemini_key", "")

    st.markdown('<span class="slabel">01 — Dossier académico</span>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Dossier", type=["pdf","txt","docx"], label_visibility="collapsed")

    st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)
    st.markdown('<span class="slabel">02 — Área o contexto del programa</span>', unsafe_allow_html=True)
    user_context = st.text_area("Contexto", label_visibility="collapsed",
        placeholder="Ej: Maestría en Administración de Hospitales, orientada a directivos de clínicas privadas en México...",
        height=100)

    st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)

    if st.button("🔍  Generar portafolio de licencias académicas", use_container_width=True):
        if not api_key:
            st.error("⚠️ Configura tu API Key de Gemini.")
        elif not uploaded and not user_context:
            st.warning("Sube el dossier o describe el contexto del programa.")
        else:
            dossier_text = extract_text(uploaded) if uploaded else ""
            with st.spinner("Analizando dossier y construyendo portafolio... esto puede tomar 30–60 segundos"):
                try:
                    result = call_gemini(api_key, dossier_text, user_context)
                except Exception as e:
                    st.error(f"Error al consultar Gemini: {e}")
                    st.stop()
            if not result.get("portafolio"):
                st.warning("No se obtuvieron resultados. Intenta con otro contexto o dossier.")
                if result.get("_raw"):
                    with st.expander("Respuesta cruda"): st.text(result["_raw"])
                st.stop()
            st.session_state["result"] = result

    st.markdown('</div>', unsafe_allow_html=True)

# ── Results ──────────────────────────────────────────────────────────
if "result" in st.session_state:
    result=st.session_state["result"]; port=result.get("portafolio",[]); conc=result.get("conceptos",[]); dorks=result.get("google_dorks",[])
    capa1=[s for s in port if s.get("capa")==1]; capa2=[s for s in port if s.get("capa")==2]; capa3=[s for s in port if s.get("capa")==3]

    st.markdown('<div class="aria-divider"></div>', unsafe_allow_html=True)

    # Stats + download
    col_s, col_d = st.columns([3,1])
    with col_s:
        st.markdown(f"""
        <div class="stats-bar">
          <div class="stat-pill">📦 <span>{len(port)}</span> herramientas</div>
          <div class="stat-pill">🔵 <span>{len(capa1)}</span> Capa 1</div>
          <div class="stat-pill">🟢 <span>{len(capa2)}</span> Capa 2</div>
          <div class="stat-pill">🟡 <span>{len(capa3)}</span> Capa 3</div>
          <div class="stat-pill">💡 <span>{len(conc)}</span> conceptos</div>
        </div>""", unsafe_allow_html=True)
    with col_d:
        st.download_button("⬇️ Descargar Excel", data=build_excel(result),
                           file_name=f"ARIA_Licencias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

    # Layer legend
    st.markdown("""
    <div class="layer-legend">
      <div class="legend-item"><div class="legend-dot" style="background:#1A7FE8"></div>Capa 1 — Core del Dossier</div>
      <div class="legend-item"><div class="legend-dot" style="background:#059669"></div>Capa 2 — Disciplinas Relacionadas</div>
      <div class="legend-item"><div class="legend-dot" style="background:#D97706"></div>Capa 3 — Ecosistema Estratégico</div>
    </div>""", unsafe_allow_html=True)

    tab1,tab2,tab3,tab4 = st.tabs([
        f"📦 Portafolio ({len(port)})", f"🗂 Paquetes",
        f"🔑 Conceptos ({len(conc)})", f"🔎 Dorks ({len(dorks)})"])

    with tab1:
        capa_labels={1:"Capa 1 — Core",2:"Capa 2 — Relacionadas",3:"Capa 3 — Ecosistema"}
        badge_cls={1:"layer-1",2:"layer-2",3:"layer-3"}
        card_cls={1:"l1",2:"l2",3:"l3"}
        for sw in port:
            capa=sw.get("capa",1); url=sw.get("url_oficial","")
            link=f'<a href="{url}" target="_blank" style="color:#1A7FE8;font-size:12px;font-weight:500;"><i class=\"ti ti-world\" aria-hidden=\"true\" style=\"font-size:13px;vertical-align:-2px;margin-right:3px;\"></i>Sitio oficial</a>' if url else ""
            s_ac=sw.get("score_academico",0); s_cv=sw.get("score_cv",0)
            s_ind=sw.get("score_industria",0); s_dif=sw.get("score_diferenciacion",0)
            score=s_ac+s_cv+s_ind+s_dif
            score_pct=int(score/20*100) if score else 0
            ac_pct=int(s_ac/5*100); cv_pct=int(s_cv/5*100); ind_pct=int(s_ind/5*100)
            bar_col="linear-gradient(90deg,#1A7FE8,#059669)" if score>=15 else ("linear-gradient(90deg,#1A7FE8,#3B82F6)" if score>=10 else "#D97706")
            acceso=sw.get("tipo_acceso","").lower()
            if any(x in acceso for x in ["grat","free","sin cost"]):
                pcls,picon,plabel="price-free","ti-circle-check","Gratis académico"
            elif any(x in acceso for x in ["descuento","discount","student","estudiante"]):
                pcls,picon,plabel="price-paid","ti-tag","Descuento académico"
            else:
                pcls,picon,plabel="price-inst","ti-building","Vía convenio institucional"
            st.markdown(f"""
            <div class="sw-card {card_cls.get(capa,'l1')}">
              <div class="sw-name">{sw.get("nombre","")}</div>
              <div style="display:flex;align-items:center;flex-wrap:wrap;gap:4px;margin-bottom:6px;">
                <span class="layer-badge {badge_cls.get(capa,'layer-1')}">{capa_labels.get(capa,"")}</span>
                <span class="price-pill {pcls}"><i class="ti {picon}" aria-hidden="true" style="font-size:11px;"></i> {plabel}</span>
              </div>
              <div class="sw-desc">{sw.get("valor_estrategico","")}</div>
              <div class="sw-detail"><strong>Acceso:</strong> {sw.get("tipo_acceso","")}</div>
              <div class="sw-detail"><strong>Módulo:</strong> {sw.get("relacion_dossier","")}</div>
              <div class="sw-detail"><strong>Certificación:</strong> {sw.get("certificacion","")}</div>
              <div class="card-divider"></div>
              <div class="score-row">
                <span class="score-label">Score total</span>
                <div class="score-track"><div class="score-fill" style="width:{score_pct}%;background:{bar_col};"></div></div>
                <span class="score-num">{score}/20</span>
              </div>
              <div style="display:flex;gap:10px;margin-top:4px;">
                <div class="score-row" style="flex:1;margin:0;">
                  <span class="score-label">Relevancia</span>
                  <div class="score-track"><div class="score-fill" style="width:{ac_pct}%;background:#1A7FE8;"></div></div>
                  <span class="score-num" style="font-size:10px;">{s_ac}/5</span>
                </div>
                <div class="score-row" style="flex:1;margin:0;">
                  <span class="score-label">Industria</span>
                  <div class="score-track"><div class="score-fill" style="width:{ind_pct}%;background:#059669;"></div></div>
                  <span class="score-num" style="font-size:10px;">{s_ind}/5</span>
                </div>
                <div class="score-row" style="flex:1;margin:0;">
                  <span class="score-label">Valor CV</span>
                  <div class="score-track"><div class="score-fill" style="width:{cv_pct}%;background:#7C3AED;"></div></div>
                  <span class="score-num" style="font-size:10px;">{s_cv}/5</span>
                </div>
              </div>
              <div style="margin-top:10px;">{link}</div>
            </div>""", unsafe_allow_html=True)

    with tab2:
        paquetes={}
        for sw in port: paquetes.setdefault(sw.get("paquete","Otros"),[]).append(sw)
        icons={"Hospitales":"🏥","Medicina y Clinica":"🩺","Cuidados Medicos":"💊","Cuidados Intensivos":"🫀",
               "Investigacion y Publicacion":"🔬","Direccion y Liderazgo":"📊","Salud Publica y Politicas":"🌎"}
        for pname in PAQUETES:
            sws=paquetes.get(pname,[])
            if not sws: continue
            st.markdown(f"""
            <div class="pkg-card">
              <div class="pkg-title">{icons.get(pname,'📦')} {pname} <span style="font-size:11px;opacity:.7;font-weight:400;">({len(sws)} herramientas)</span></div>
              <div class="pkg-obj">Herramientas de nivel institucional en {pname.lower()}.</div>
              <div class="pkg-orgs">{" · ".join(s.get("nombre","") for s in sws)}</div>
              <div class="pkg-benefit">✓ Fortalece CV y empleabilidad en el sector salud.</div>
            </div>""", unsafe_allow_html=True)

    with tab3:
        for concept in conc:
            with st.expander(f"🔑 {concept.get('palabra_clave','')}"):
                st.markdown(f"**Tecnologías:** {concept.get('tecnologias_asociadas','')}")
                st.markdown(f"**Métodos / Marcos:** {concept.get('metodos_marcos','')}")
                st.markdown(f"**Áreas de aplicación:** {concept.get('areas_aplicacion','')}")

    with tab4:
        st.caption("Copia cualquier dork y pégalo directamente en Google")
        for i,dork in enumerate(dorks,1):
            st.markdown(f"""
            <div class="dork-card">
              <span class="dork-num">{i}</span>
              <span class="dork-text">{dork}</span>
            </div>""", unsafe_allow_html=True)

    st.markdown('<p style="text-align:center;font-size:11px;color:#94A3B8;margin-top:40px;">ARIA Licencias · Powered by Google Gemini 2.5 Flash</p>', unsafe_allow_html=True)
