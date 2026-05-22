import streamlit as st
import google.generativeai as genai
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
import json, re, io
from datetime import datetime

st.set_page_config(page_title="ARIA Buscador", page_icon="🎓", layout="centered", initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Mono:wght@400;500&family=Sora:wght@300;400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'Sora', sans-serif; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 820px; }
.aria-header { background: linear-gradient(135deg, #0D1B2A 0%, #1A3C5E 100%); border-radius: 16px; padding: 28px 36px; text-align: center; margin-bottom: 24px; border: 1px solid rgba(201,168,76,0.2); position: relative; overflow: hidden; }
.aria-header::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, #1A3C5E, #1A6FA8, #C9A84C); }
.aria-title { font-family: 'DM Serif Display', serif; font-size: 2.4rem; color: #fff; margin: 0; }
.aria-title em { font-style: italic; color: #E8C96A; }
.aria-sub { font-size: 0.82rem; color: rgba(255,255,255,0.5); margin-top: 6px; }
.section-label { font-size: 0.68rem; letter-spacing: 2.5px; text-transform: uppercase; color: #8A9BAE; font-weight: 600; margin-bottom: 4px; }
.layer-badge { display: inline-block; border-radius: 20px; padding: 2px 10px; font-size: 0.68rem; font-weight: 600; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 8px; }
.layer-1 { background: #EFF6FF; color: #1D4ED8; } .layer-2 { background: #F0FDF4; color: #166534; } .layer-3 { background: #FFF7ED; color: #9A3412; }
.sw-card { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 14px; padding: 18px 20px; margin-bottom: 12px; }
.sw-card.l1 { border-left: 4px solid #1D4ED8; } .sw-card.l2 { border-left: 4px solid #166534; } .sw-card.l3 { border-left: 4px solid #EA580C; }
.sw-name { font-size: 1rem; font-weight: 600; color: #0D1B2A; margin-bottom: 3px; }
.sw-desc { font-size: 0.82rem; color: #475569; line-height: 1.5; margin-bottom: 6px; }
.sw-detail { font-size: 0.78rem; color: #64748b; margin-bottom: 2px; }
.sw-detail strong { color: #334155; }
.pkg-card { background: linear-gradient(135deg, #0D1B2A, #1A3C5E); border-radius: 14px; padding: 18px 20px; margin-bottom: 12px; color: white; }
.pkg-title { font-size: 1rem; font-weight: 600; color: #E8C96A; margin-bottom: 4px; }
.pkg-obj { font-size: 0.8rem; color: rgba(255,255,255,0.65); margin-bottom: 8px; }
.pkg-sw { font-size: 0.8rem; color: rgba(255,255,255,0.85); }
.pkg-benefit { font-size: 0.78rem; color: #86EFAC; margin-top: 6px; }
.dork-card { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 10px 14px; margin-bottom: 8px; display: flex; align-items: flex-start; gap: 10px; }
.dork-num { font-size: 0.75rem; color: #94a3b8; font-weight: 600; min-width: 20px; }
.dork-text { font-family: 'DM Mono', monospace; font-size: 0.78rem; color: #166534; line-height: 1.5; flex: 1; word-break: break-all; }
.stats-bar { background: #0D1B2A; border-radius: 10px; padding: 12px 20px; display: flex; gap: 20px; align-items: center; margin-bottom: 20px; flex-wrap: wrap; }
.stat-item { font-size: 0.82rem; color: rgba(255,255,255,0.7); }
.stat-item span { color: #E8C96A; font-weight: 600; }
.aria-divider { height: 1px; background: #e2e8f0; margin: 24px 0; }
.layer-legend { display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }
.legend-item { display: flex; align-items: center; gap: 6px; font-size: 0.78rem; color: #475569; }
.legend-dot { width: 10px; height: 10px; border-radius: 50%; }
.stButton > button { background: linear-gradient(135deg, #1A3C5E, #1A6FA8) !important; color: white !important; border: none !important; border-radius: 10px !important; padding: 0.6rem 2rem !important; font-family: 'Sora', sans-serif !important; font-weight: 500 !important; font-size: 0.95rem !important; width: 100% !important; }
.stTextArea textarea { border-radius: 10px !important; border: 1px solid #cbd5e1 !important; }
.stFileUploader { border: 1.5px dashed #94a3b8 !important; border-radius: 12px !important; padding: 8px !important; }
.stDownloadButton > button { background: #166534 !important; color: white !important; border: none !important; border-radius: 10px !important; font-family: 'Sora', sans-serif !important; font-weight: 500 !important; width: 100% !important; }
</style>
""", unsafe_allow_html=True)


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
    msg = f"DOSSIER:\n{dossier[:6000] if dossier else 'Sin dossier.'}\n\nCONTEXTO:\n{context or 'Posgrado en ciencias de la salud.'}\n\nResponde SOLO con el JSON."
    raw = model.generate_content(msg).text
    result = parse_json(raw)
    if result:
        return result
    return {"portafolio":[], "conceptos":[], "google_dorks":[], "_raw": raw[:800]}


CAPA_COLORS = {1:"1D4ED8", 2:"166534", 3:"EA580C"}
CAPA_BG     = {1:"DBEAFE", 2:"DCFCE7", 3:"FFEDD5"}
PAQUETES    = ["Hospitales","Medicina y Clinica","Cuidados Medicos","Cuidados Intensivos",
               "Investigacion y Publicacion","Direccion y Liderazgo","Salud Publica y Politicas"]

def _h(ws,r,c,v,bg="1A3C5E",fg="FFFFFF",sz=11):
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
    ws1.merge_cells("A1:I1"); tc=ws1.cell(1,1,"PORTAFOLIO ESTRATEGICO — ARIA Buscador")
    tc.font=Font(bold=True,color="FFFFFF",name="Arial",size=13); tc.fill=PatternFill("solid",fgColor="0D1B2A")
    tc.alignment=Alignment(horizontal="center",vertical="center"); ws1.row_dimensions[1].height=32
    for i,(h,w) in enumerate(zip(["#","Nombre","Tipo","Capa","Acceso","Link","Relacion Dossier","Valor Estrategico","Certificacion"],[4,24,18,8,24,32,36,36,30]),1):
        _h(ws1,2,i,h); ws1.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    ws1.row_dimensions[2].height=30
    for idx,sw in enumerate(port,1):
        row=idx+2; capa=sw.get("capa",1); alt="F8FAFC" if idx%2==0 else "FFFFFF"
        for col,val in enumerate([idx,sw.get("nombre",""),"Premium Institucional",f"Capa {capa}",sw.get("tipo_acceso",""),sw.get("url_oficial",""),sw.get("relacion_dossier",""),sw.get("valor_estrategico",""),sw.get("certificacion","")],1):
            _c(ws1,row,col,val,bold=(col==4),color=CAPA_COLORS.get(capa,"334155") if col==4 else ("1A6FA8" if col==6 else "334155"),bg=CAPA_BG.get(capa,alt) if col==4 else alt,ha="center" if col in(1,3,4) else "left")
        ws1.row_dimensions[row].height=52
    ws1.freeze_panes="A3"; ws1.auto_filter.ref=f"A2:I{len(port)+2}"

    ws2=wb.create_sheet("Paquetes"); ws2.merge_cells("A1:D1")
    tc2=ws2.cell(1,1,"PAQUETES ESTRATEGICOS"); tc2.font=Font(bold=True,color="FFFFFF",name="Arial",size=13)
    tc2.fill=PatternFill("solid",fgColor="0D1B2A"); tc2.alignment=Alignment(horizontal="center",vertical="center"); ws2.row_dimensions[1].height=32
    for i,(h,w) in enumerate(zip(["Paquete","Objetivo","Software","Beneficio"],[24,40,50,40]),1):
        _h(ws2,2,i,h); ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    paquetes={}
    for sw in port: paquetes.setdefault(sw.get("paquete","Otros"),[]).append(sw)
    pkg_bgs=["0D1B2A","1A3C5E","1A6FA8","166534","7C3AED","9A3412","065F46"]; row2=3
    for pi,pname in enumerate(PAQUETES):
        sws=paquetes.get(pname,[]); 
        if not sws: continue
        bg_col=pkg_bgs[pi%len(pkg_bgs)]; names=", ".join(s.get("nombre","") for s in sws[:12])
        for col,val in enumerate([pname,f"Herramientas de {pname} nivel institucional.",names,f"Fortalece CV en {pname}."],1):
            c=ws2.cell(row2,col,val); c.font=Font(bold=(col==1),color="FFFFFF" if col==1 else "0D1B2A",name="Arial",size=10)
            c.fill=PatternFill("solid",fgColor=bg_col if col==1 else ("F8FAFC" if row2%2==0 else "FFFFFF"))
            c.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True)
            thin=Side(style="thin",color="E2E8F0"); c.border=Border(left=thin,right=thin,top=thin,bottom=thin)
        ws2.row_dimensions[row2].height=60; row2+=1
    ws2.freeze_panes="A3"

    ws3=wb.create_sheet("Conceptos"); ws3.merge_cells("A1:D1")
    tc3=ws3.cell(1,1,"CONCEPTOS CLAVE"); tc3.font=Font(bold=True,color="FFFFFF",name="Arial",size=13)
    tc3.fill=PatternFill("solid",fgColor="0D1B2A"); tc3.alignment=Alignment(horizontal="center",vertical="center"); ws3.row_dimensions[1].height=32
    for i,(h,w) in enumerate(zip(["Palabra Clave","Tecnologias","Metodos","Areas"],[36,40,36,36]),1):
        _h(ws3,2,i,h,bg="1A3C5E"); ws3.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    for idx,c in enumerate(conc,1):
        row3=idx+2; bg3="EFF6FF" if idx%2==0 else "FFFFFF"
        for col,key in enumerate(["palabra_clave","tecnologias_asociadas","metodos_marcos","areas_aplicacion"],1):
            _c(ws3,row3,col,c.get(key,""),bg=bg3)
        ws3.row_dimensions[row3].height=40
    ws3.freeze_panes="A3"

    ws4=wb.create_sheet("Scoring"); ws4.merge_cells("A1:G1")
    tc4=ws4.cell(1,1,"SCORING ESTRATEGICO"); tc4.font=Font(bold=True,color="FFFFFF",name="Arial",size=13)
    tc4.fill=PatternFill("solid",fgColor="0D1B2A"); tc4.alignment=Alignment(horizontal="center",vertical="center"); ws4.row_dimensions[1].height=32
    for i,(h,w) in enumerate(zip(["#","Nombre","Relevancia","Valor CV","Industria","Diferenciacion","SCORE TOTAL"],[4,30,18,14,16,20,14]),1):
        _h(ws4,2,i,h,bg="1A3C5E"); ws4.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    ws4.row_dimensions[2].height=36
    for idx,sw in enumerate(port,1):
        row4=idx+2; bg4="F8FAFC" if idx%2==0 else "FFFFFF"; capa=sw.get("capa",1)
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
    tc5=ws5.cell(1,1,"GOOGLE DORKS"); tc5.font=Font(bold=True,color="FFFFFF",name="Arial",size=13)
    tc5.fill=PatternFill("solid",fgColor="166534"); tc5.alignment=Alignment(horizontal="center",vertical="center"); ws5.row_dimensions[1].height=32
    for i,(h,w) in enumerate(zip(["#","Dork","Uso"],[6,80,40]),1):
        _h(ws5,2,i,h,bg="166534"); ws5.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    for idx,dork in enumerate(dorks,1):
        row5=idx+2; bg5="F0FDF4" if idx%2==0 else "FFFFFF"
        _c(ws5,row5,1,idx,ha="center",bg=bg5)
        c=ws5.cell(row5,2,dork); c.font=Font(name="Courier New",size=10,color="166534")
        c.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True)
        thin=Side(style="thin",color="E2E8F0"); c.border=Border(left=thin,right=thin,top=thin,bottom=thin); c.fill=PatternFill("solid",fgColor=bg5)
        _c(ws5,row5,3,"",bg=bg5); ws5.row_dimensions[row5].height=22
    ws5.freeze_panes="A3"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf.getvalue()


# ── UI ───────────────────────────────────────────────────────────────
st.markdown('<div class="aria-header"><h1 class="aria-title">ARIA <em>Buscador</em></h1><p class="aria-sub">Portafolio estratégico de software para programas académicos de posgrado</p></div>', unsafe_allow_html=True)

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

st.markdown('<div class="section-label">01 — Dossier académico (PDF, TXT o DOCX)</div>', unsafe_allow_html=True)
uploaded = st.file_uploader("Dossier", type=["pdf","txt","docx"], label_visibility="collapsed")
st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)
st.markdown('<div class="section-label">02 — Contexto adicional (opcional)</div>', unsafe_allow_html=True)
user_context = st.text_area("Contexto", label_visibility="collapsed", placeholder="Ej: Maestría en Administración de Hospitales...", height=90)

if st.button("🔍  Generar portafolio estratégico", use_container_width=True):
    if not api_key:
        st.error("⚠️ Configura tu API Key de Gemini.")
    elif not uploaded and not user_context:
        st.warning("Sube el dossier o describe el contexto.")
    else:
        dossier_text = extract_text(uploaded) if uploaded else ""
        with st.spinner("Analizando dossier y construyendo portafolio... (30–60 segundos)"):
            try:
                result = call_gemini(api_key, dossier_text, user_context)
            except Exception as e:
                st.error(f"Error al consultar Gemini: {e}")
                st.stop()
        if not result.get("portafolio"):
            st.warning("No se obtuvieron resultados.")
            if result.get("_raw"):
                with st.expander("Respuesta cruda"): st.text(result["_raw"])
            st.stop()
        st.session_state["result"] = result

if "result" in st.session_state:
    result=st.session_state["result"]; port=result.get("portafolio",[]); conc=result.get("conceptos",[]); dorks=result.get("google_dorks",[])
    capa1=[s for s in port if s.get("capa")==1]; capa2=[s for s in port if s.get("capa")==2]; capa3=[s for s in port if s.get("capa")==3]
    st.markdown('<div class="aria-divider"></div>', unsafe_allow_html=True)
    col_s,col_d=st.columns([2,1])
    with col_s:
        st.markdown(f'<div class="stats-bar"><div class="stat-item"><span>{len(port)}</span> herramientas</div><div class="stat-item"><span>{len(capa1)}</span> Capa 1</div><div class="stat-item"><span>{len(capa2)}</span> Capa 2</div><div class="stat-item"><span>{len(capa3)}</span> Capa 3</div><div class="stat-item"><span>{len(conc)}</span> conceptos</div></div>', unsafe_allow_html=True)
    with col_d:
        st.download_button("⬇️ Descargar Excel (5 hojas)", data=build_excel(result), file_name=f"ARIA_Portafolio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    st.markdown('<div class="layer-legend"><div class="legend-item"><div class="legend-dot" style="background:#1D4ED8"></div>Capa 1 — Core</div><div class="legend-item"><div class="legend-dot" style="background:#166534"></div>Capa 2 — Relacionadas</div><div class="legend-item"><div class="legend-dot" style="background:#EA580C"></div>Capa 3 — Ecosistema</div></div>', unsafe_allow_html=True)

    tab1,tab2,tab3,tab4=st.tabs([f"📦 Portafolio ({len(port)})",f"🗂 Paquetes",f"🔑 Conceptos ({len(conc)})",f"🔎 Dorks ({len(dorks)})"])

    with tab1:
        capa_labels={1:"Capa 1 — Core",2:"Capa 2 — Relacionadas",3:"Capa 3 — Ecosistema"}
        badge_cls={1:"layer-1",2:"layer-2",3:"layer-3"}; card_cls={1:"l1",2:"l2",3:"l3"}
        for sw in port:
            capa=sw.get("capa",1); url=sw.get("url_oficial","")
            link=f'<a href="{url}" target="_blank" style="color:#1A6FA8;font-size:0.78rem;">🌐 Sitio oficial</a>' if url else ""
            score=sw.get("score_academico",0)+sw.get("score_cv",0)+sw.get("score_industria",0)+sw.get("score_diferenciacion",0)
            st.markdown(f'<div class="sw-card {card_cls.get(capa,"l1")}"><div class="sw-name">{sw.get("nombre","")}</div><span class="layer-badge {badge_cls.get(capa,"layer-1")}">{capa_labels.get(capa,"")}</span>&nbsp;<span style="font-size:0.75rem;color:#64748b;">Score: <strong>{score}/20</strong></span><div class="sw-desc">{sw.get("valor_estrategico","")}</div><div class="sw-detail"><strong>Acceso:</strong> {sw.get("tipo_acceso","")}</div><div class="sw-detail"><strong>Módulo:</strong> {sw.get("relacion_dossier","")}</div><div class="sw-detail"><strong>Certificación:</strong> {sw.get("certificacion","")}</div><div style="margin-top:8px;">{link}</div></div>', unsafe_allow_html=True)

    with tab2:
        paquetes={}
        for sw in port: paquetes.setdefault(sw.get("paquete","Otros"),[]).append(sw)
        icons={"Hospitales":"🏥","Medicina y Clinica":"🩺","Cuidados Medicos":"💊","Cuidados Intensivos":"🫀","Investigacion y Publicacion":"🔬","Direccion y Liderazgo":"📊","Salud Publica y Politicas":"🌎"}
        for pname in PAQUETES:
            sws=paquetes.get(pname,[])
            if not sws: continue
            st.markdown(f'<div class="pkg-card"><div class="pkg-title">{icons.get(pname,"📦")} {pname} <span style="font-size:0.75rem;opacity:.7;">({len(sws)} herramientas)</span></div><div class="pkg-obj">Herramientas de nivel institucional en {pname.lower()}.</div><div class="pkg-sw">{" · ".join(s.get("nombre","") for s in sws)}</div><div class="pkg-benefit">✓ Fortalece CV y empleabilidad en el sector salud.</div></div>', unsafe_allow_html=True)

    with tab3:
        for concept in conc:
            with st.expander(f"🔑 {concept.get('palabra_clave','')}"):
                st.markdown(f"**Tecnologías:** {concept.get('tecnologias_asociadas','')}")
                st.markdown(f"**Métodos:** {concept.get('metodos_marcos','')}")
                st.markdown(f"**Áreas:** {concept.get('areas_aplicacion','')}")

    with tab4:
        st.caption("Copia cualquier dork y pégalo en Google")
        for i,dork in enumerate(dorks,1):
            st.markdown(f'<div class="dork-card"><span class="dork-num">{i}</span><span class="dork-text">{dork}</span></div>', unsafe_allow_html=True)

    st.markdown('<p style="text-align:center;font-size:0.72rem;color:#94a3b8;margin-top:32px;">ARIA Buscador · Powered by Google Gemini 2.5 Flash</p>', unsafe_allow_html=True)
