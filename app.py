import streamlit as st
import google.generativeai as genai
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
import json, re, io, os
from datetime import datetime

# ── Page config ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="ARIA Buscador",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Mono:wght@400;500&family=Sora:wght@300;400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'Sora', sans-serif; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 820px; }

.aria-header {
    background: linear-gradient(135deg, #0D1B2A 0%, #1A3C5E 100%);
    border-radius: 16px; padding: 28px 36px; text-align: center;
    margin-bottom: 24px; border: 1px solid rgba(201,168,76,0.2);
    position: relative; overflow: hidden;
}
.aria-header::before {
    content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px;
    background: linear-gradient(90deg, #1A3C5E, #1A6FA8, #C9A84C);
}
.aria-title { font-family: 'DM Serif Display', serif; font-size: 2.4rem; color: #fff; margin: 0; }
.aria-title em { font-style: italic; color: #E8C96A; }
.aria-sub { font-size: 0.82rem; color: rgba(255,255,255,0.5); margin-top: 6px; }

.section-label {
    font-size: 0.68rem; letter-spacing: 2.5px; text-transform: uppercase;
    color: #8A9BAE; font-weight: 600; margin-bottom: 4px;
}

.layer-badge {
    display: inline-block; border-radius: 20px; padding: 2px 10px;
    font-size: 0.68rem; font-weight: 600; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 8px;
}
.layer-1 { background: #EFF6FF; color: #1D4ED8; }
.layer-2 { background: #F0FDF4; color: #166534; }
.layer-3 { background: #FFF7ED; color: #9A3412; }

.sw-card {
    background: #f8fafc; border: 1px solid #e2e8f0;
    border-radius: 14px; padding: 18px 20px; margin-bottom: 12px;
}
.sw-card.l1 { border-left: 4px solid #1D4ED8; }
.sw-card.l2 { border-left: 4px solid #166534; }
.sw-card.l3 { border-left: 4px solid #EA580C; }
.sw-name { font-size: 1rem; font-weight: 600; color: #0D1B2A; margin-bottom: 3px; }
.sw-desc { font-size: 0.82rem; color: #475569; line-height: 1.5; margin-bottom: 6px; }
.sw-detail { font-size: 0.78rem; color: #64748b; margin-bottom: 2px; }
.sw-detail strong { color: #334155; }
.sw-verified { color: #16a34a; font-size: 0.75rem; font-weight: 600; margin-top: 4px; }

.pkg-card {
    background: linear-gradient(135deg, #0D1B2A, #1A3C5E);
    border-radius: 14px; padding: 18px 20px; margin-bottom: 12px; color: white;
}
.pkg-title { font-size: 1rem; font-weight: 600; color: #E8C96A; margin-bottom: 4px; }
.pkg-obj { font-size: 0.8rem; color: rgba(255,255,255,0.65); margin-bottom: 8px; }
.pkg-sw { font-size: 0.8rem; color: rgba(255,255,255,0.85); }
.pkg-benefit { font-size: 0.78rem; color: #86EFAC; margin-top: 6px; }

.dork-card {
    background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px;
    padding: 10px 14px; margin-bottom: 8px; display: flex; align-items: flex-start; gap: 10px;
}
.dork-num { font-size: 0.75rem; color: #94a3b8; font-weight: 600; min-width: 20px; }
.dork-text { font-family: 'DM Mono', monospace; font-size: 0.78rem; color: #166534; line-height: 1.5; flex: 1; word-break: break-all; }

.stats-bar {
    background: #0D1B2A; border-radius: 10px; padding: 12px 20px;
    display: flex; gap: 20px; align-items: center; margin-bottom: 20px; flex-wrap: wrap;
}
.stat-item { font-size: 0.82rem; color: rgba(255,255,255,0.7); }
.stat-item span { color: #E8C96A; font-weight: 600; }

.aria-divider { height: 1px; background: #e2e8f0; margin: 24px 0; }
.layer-legend { display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }
.legend-item { display: flex; align-items: center; gap: 6px; font-size: 0.78rem; color: #475569; }
.legend-dot { width: 10px; height: 10px; border-radius: 50%; }

.stButton > button {
    background: linear-gradient(135deg, #1A3C5E, #1A6FA8) !important;
    color: white !important; border: none !important; border-radius: 10px !important;
    padding: 0.6rem 2rem !important; font-family: 'Sora', sans-serif !important;
    font-weight: 500 !important; font-size: 0.95rem !important; width: 100% !important;
}
.stButton > button:hover { opacity: 0.88 !important; }
.stTextArea textarea {
    border-radius: 10px !important; border: 1px solid #cbd5e1 !important;
    font-family: 'Sora', sans-serif !important; font-size: 0.9rem !important;
}
.stFileUploader { border: 1.5px dashed #94a3b8 !important; border-radius: 12px !important; padding: 8px !important; }
.stTextInput input { border-radius: 10px !important; border: 1px solid #cbd5e1 !important; font-family: 'Sora', sans-serif !important; }
.stDownloadButton > button {
    background: #166534 !important; color: white !important; border: none !important;
    border-radius: 10px !important; font-family: 'Sora', sans-serif !important;
    font-weight: 500 !important; width: 100% !important;
}
</style>
""", unsafe_allow_html=True)


# ── Helpers ──────────────────────────────────────────────────────────
def extract_text(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.txt'):
        return uploaded_file.read().decode('utf-8', errors='ignore')
    elif name.endswith('.pdf'):
        try:
            import pypdf
            reader = pypdf.PdfReader(uploaded_file)
            return '\n'.join(p.extract_text() or '' for p in reader.pages)
        except Exception:
            return ""
    elif name.endswith('.docx'):
        try:
            import docx
            doc = docx.Document(io.BytesIO(uploaded_file.read()))
            return '\n'.join(p.text for p in doc.paragraphs)
        except Exception:
            return ""
    return ""


# ── SYSTEM PROMPT (oculto al usuario) ────────────────────────────────
SYSTEM_PROMPT = """Actúa como un consultor senior en estrategia (nivel McKinsey) especializado en:
- Vinculación académica internacional
- Desarrollo de portafolios tecnológicos universitarios
- Posicionamiento premium de programas educativos de posgrado en ciencias de la salud

El usuario te compartirá un dossier académico de posgrado.
Tu tarea es diseñar un portafolio estratégico de software y herramientas tecnológicas,
alineado al dossier, bajo un enfoque de valor percibido premium.
NO es un catálogo técnico — es una arquitectura de posicionamiento institucional.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OBJETIVO
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Construir un portafolio que:
- Eleve el posicionamiento de la universidad frente a programas competidores
- Aumente la percepción de valor del programa ante estudiantes y acreditadores
- Conecte al estudiante con herramientas reales de la industria hospitalaria
- Sea comercialmente defendible ante organismos externos

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
MODELO DE TRES CAPAS (OBLIGATORIO)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Toda herramienta debe clasificarse en exactamente una capa:

CAPA 1 — CORE DEL DOSSIER
  Herramientas directamente alineadas con la disciplina central.
  Ejemplo para Enfermería: CINAHL, NursingCentral, Nursys e-Notify

CAPA 2 — DISCIPLINAS RELACIONADAS
  Sistemas de soporte a la dirección: calidad, finanzas sanitarias,
  RRHH, estadística, investigación, acreditación hospitalaria.
  Ejemplo: SAP Healthcare, ISOTools, SPSS, UpToDate

CAPA 3 — ECOSISTEMA ESTRATÉGICO
  Gestión hospitalaria integral, simulación clínica, liderazgo ejecutivo,
  mercadotecnia sanitaria, salud pública, investigación científica.
  Ejemplo: Epic EHR, Salesforce Health Cloud, Laerdal SimMan

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FILTROS DE INCLUSIÓN (todos obligatorios)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Solo incluir herramientas que cumplan TODOS estos criterios:
✔ Licencia académica o institucional (no acceso público general)
✔ Acceso restringido mediante: afiliación universitaria, correo institucional, convenio formal o consorcio académico
✔ Uso real y verificable en academia o industria del sector salud
✔ Certificación oficial, learning path estructurado o uso guiado en laboratorio o investigación
✔ Reconocimiento internacional verificable

FILTROS DE EXCLUSIÓN (cualquiera descalifica):
✗ Software completamente gratuito sin control de acceso
✗ Trials o versiones de prueba con vigencia menor a 12 meses
✗ Herramientas sin uso real demostrable en academia o industria
✗ Plataformas freemium sin restricción académica efectiva
✗ Software poco reconocido o sin respaldo institucional

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CONDICIONES NO NEGOCIABLES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- Mínimo 50 herramientas — sin duplicados
- Nunca inventar software; solo herramientas con URL oficial verificable
- Cada herramienta debe mapear a al menos un módulo del plan de estudios
- Juicio crítico sobre cada herramienta — no descripciones promocionales

Responde ÚNICAMENTE con un JSON válido con esta estructura exacta (sin texto adicional, sin bloques de código):
{
  "portafolio": [
    {
      "nombre": "string",
      "capa": 1,
      "tipo_acceso": "string",
      "url_oficial": "string",
      "relacion_dossier": "string (máx 25 palabras, citar módulo)",
      "valor_estrategico": "string",
      "certificacion": "string",
      "paquete": "string (uno de: Hospitales | Medicina y Clínica | Cuidados Médicos | Cuidados Intensivos | Investigación y Publicación | Dirección y Liderazgo | Salud Pública y Políticas)",
      "score_academico": 1,
      "score_cv": 1,
      "score_industria": 1,
      "score_diferenciacion": 1
    }
  ],
  "conceptos": [
    {
      "palabra_clave": "string",
      "tecnologias_asociadas": "string",
      "metodos_marcos": "string",
      "areas_aplicacion": "string"
    }
  ],
  "google_dorks": ["dork1", "dork2"]
}

Los scores son enteros del 1 al 5. Genera al menos 50 herramientas en portafolio y al menos 10 conceptos clave."""


def search_with_gemini(api_key: str, dossier_text: str, user_context: str) -> dict:
    genai.configure(api_key=api_key)

    user_message = f"""Analiza el siguiente dossier académico y construye el portafolio estratégico completo.

DOSSIER ACADÉMICO:
{dossier_text[:6000] if dossier_text else 'No se proporcionó dossier. Usa el contexto adicional.'}

CONTEXTO ADICIONAL DEL USUARIO:
{user_context or 'Programa de posgrado en ciencias de la salud.'}

Busca activamente en internet para verificar cada herramienta antes de incluirla.
Responde SOLO con el JSON."""

    model = genai.GenerativeModel(
        model_name="gemini-2.0-flash",
    
        system_instruction=SYSTEM_PROMPT,
    )

    response = model.generate_content(user_message)
    raw = response.text

    try:
        clean = re.sub(r'```json|```', '', raw).strip()
        match = re.search(r'\{[\s\S]*\}', clean)
        if match:
            return json.loads(match.group())
    except Exception:
        pass
    return {"portafolio": [], "conceptos": [], "google_dorks": [], "_raw": raw[:800]}


# ── Excel builder ─────────────────────────────────────────────────────
CAPA_COLORS = {1: "1D4ED8", 2: "166534", 3: "EA580C"}
CAPA_BG     = {1: "DBEAFE", 2: "DCFCE7", 3: "FFEDD5"}
PAQUETES_ORDER = ["Hospitales","Medicina y Clínica","Cuidados Médicos",
                  "Cuidados Intensivos","Investigación y Publicación",
                  "Dirección y Liderazgo","Salud Pública y Políticas"]

def _hdr(ws, row, col, value, bg="1A3C5E", fg="FFFFFF", sz=11, bold=True):
    c = ws.cell(row, col, value)
    c.font      = Font(bold=bold, color=fg, name="Arial", size=sz)
    c.fill      = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin        = Side(style="thin", color="CCCCCC")
    c.border    = Border(left=thin, right=thin, top=thin, bottom=thin)
    return c

def _cell(ws, row, col, value, bold=False, color="000000", bg=None, wrap=True, halign="left"):
    c = ws.cell(row, col, value)
    c.font      = Font(bold=bold, color=color, name="Arial", size=10)
    c.alignment = Alignment(horizontal=halign, vertical="center", wrap_text=wrap)
    thin        = Side(style="thin", color="E2E8F0")
    c.border    = Border(left=thin, right=thin, top=thin, bottom=thin)
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    return c

def build_excel(data: dict) -> bytes:
    wb   = openpyxl.Workbook()
    port = data.get("portafolio", [])
    conc = data.get("conceptos", [])
    dorks= data.get("google_dorks", [])

    # ── HOJA 1: PORTAFOLIO ESTRATÉGICO ──────────────────────────────
    ws1 = wb.active
    ws1.title = "Portafolio Estratégico"

    h1_cols = ["#","Nombre del Software","Tipo","Capa","Tipo de Acceso",
               "Link Oficial","Relación con el Dossier","Valor Estratégico","Certificación / Uso Guiado"]
    h1_widths= [4, 24, 18, 8, 24, 32, 36, 36, 30]

    # Title row
    ws1.merge_cells("A1:I1")
    tc = ws1.cell(1,1,"PORTAFOLIO ESTRATÉGICO DE SOFTWARE — ARIA Buscador")
    tc.font      = Font(bold=True, color="FFFFFF", name="Arial", size=13)
    tc.fill      = PatternFill("solid", fgColor="0D1B2A")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 32

    for i,(h,w) in enumerate(zip(h1_cols, h1_widths), 1):
        _hdr(ws1, 2, i, h)
        ws1.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws1.row_dimensions[2].height = 32

    for idx, sw in enumerate(port, 1):
        row = idx + 2
        capa = sw.get("capa", 1)
        bg   = CAPA_BG.get(capa, "FFFFFF")
        fg   = CAPA_COLORS.get(capa, "000000")
        alt  = "F8FAFC" if idx % 2 == 0 else "FFFFFF"
        row_bg = bg if True else alt

        vals = [
            idx,
            sw.get("nombre",""),
            "Premium Institucional",
            f"Capa {capa}",
            sw.get("tipo_acceso",""),
            sw.get("url_oficial",""),
            sw.get("relacion_dossier",""),
            sw.get("valor_estrategico",""),
            sw.get("certificacion",""),
        ]
        for col, val in enumerate(vals, 1):
            is_link = col == 6
            cell_bg = bg if col == 4 else ("F8FAFC" if idx%2==0 else "FFFFFF")
            c = _cell(ws1, row, col, val, bg=cell_bg,
                      color=fg if col==4 else ("1A6FA8" if is_link else "334155"),
                      bold=(col==4), halign="center" if col in(1,3,4) else "left")
        ws1.row_dimensions[row].height = 52

    ws1.freeze_panes = "A3"
    ws1.auto_filter.ref = f"A2:I{len(port)+2}"

    # Leyenda de capas
    leg_row = len(port) + 4
    ws1.merge_cells(f"A{leg_row}:B{leg_row}")
    ws1.cell(leg_row,1,"LEYENDA DE CAPAS").font = Font(bold=True, name="Arial", size=10)
    for capa, label, desc in [
        (1,"Capa 1 — Core del Dossier","Herramientas directamente alineadas a la disciplina central"),
        (2,"Capa 2 — Disciplinas Relacionadas","Calidad, finanzas sanitarias, RRHH, estadística, investigación"),
        (3,"Capa 3 — Ecosistema Estratégico","Gestión hospitalaria, simulación clínica, liderazgo ejecutivo"),
    ]:
        leg_row += 1
        c1 = ws1.cell(leg_row, 1, label)
        c1.font = Font(bold=True, color=CAPA_COLORS[capa], name="Arial", size=10)
        c1.fill = PatternFill("solid", fgColor=CAPA_BG[capa])
        c2 = ws1.cell(leg_row, 2, desc)
        c2.font = Font(name="Arial", size=10)

    # ── HOJA 2: PAQUETES ESTRATÉGICOS ───────────────────────────────
    ws2 = wb.create_sheet("Paquetes Estratégicos")
    ws2.merge_cells("A1:F1")
    tc2 = ws2.cell(1,1,"PAQUETES ESTRATÉGICOS POR CONTEXTO DE USO")
    tc2.font      = Font(bold=True, color="FFFFFF", name="Arial", size=13)
    tc2.fill      = PatternFill("solid", fgColor="0D1B2A")
    tc2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 32

    pkg_headers = ["Paquete","Objetivo Estratégico","Software Incluido","Beneficio para el Estudiante"]
    pkg_widths  = [24, 40, 50, 40]
    for i,(h,w) in enumerate(zip(pkg_headers, pkg_widths),1):
        _hdr(ws2, 2, i, h)
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws2.row_dimensions[2].height = 28

    # Group software by paquete
    paquetes: dict = {}
    for sw in port:
        pkg = sw.get("paquete","Otros")
        paquetes.setdefault(pkg, []).append(sw)

    pkg_colors = ["0D1B2A","1A3C5E","1A6FA8","166534","7C3AED","9A3412","065F46"]
    row = 3
    for pi, pname in enumerate(PAQUETES_ORDER):
        sws = paquetes.get(pname, [])
        if not sws:
            continue
        bg_col = pkg_colors[pi % len(pkg_colors)]
        obj = f"Equipar al estudiante con herramientas de {pname.lower()} de nivel institucional."
        sw_names = ", ".join(s.get("nombre","") for s in sws[:12])
        benefit  = f"Acceso a stack real de {pname.lower()}; fortalece CV y empleabilidad en el sector."

        for col, (val, w) in enumerate(zip(
            [pname, obj, sw_names, benefit],
            pkg_widths
        ), 1):
            c = ws2.cell(row, col, val)
            c.font      = Font(bold=(col==1), color="FFFFFF" if col==1 else "0D1B2A", name="Arial", size=10)
            c.fill      = PatternFill("solid", fgColor=bg_col if col==1 else ("F8FAFC" if row%2==0 else "FFFFFF"))
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            thin        = Side(style="thin", color="E2E8F0")
            c.border    = Border(left=thin,right=thin,top=thin,bottom=thin)
        ws2.row_dimensions[row].height = 60
        row += 1

    ws2.freeze_panes = "A3"

    # ── HOJA 3: CONCEPTOS DEL DOSSIER ───────────────────────────────
    ws3 = wb.create_sheet("Conceptos del Dossier")
    ws3.merge_cells("A1:D1")
    tc3 = ws3.cell(1,1,"CONCEPTOS CLAVE DEL DOSSIER — MAPA TECNOLÓGICO")
    tc3.font      = Font(bold=True, color="FFFFFF", name="Arial", size=13)
    tc3.fill      = PatternFill("solid", fgColor="0D1B2A")
    tc3.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 32

    c3_headers = ["Palabras Clave del Plan de Estudios","Tecnologías y Herramientas Asociadas",
                  "Métodos y Marcos de Referencia","Áreas de Aplicación Profesional"]
    c3_widths  = [36, 40, 36, 36]
    for i,(h,w) in enumerate(zip(c3_headers, c3_widths),1):
        _hdr(ws3, 2, i, h, bg="1A3C5E")
        ws3.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws3.row_dimensions[2].height = 32

    for idx, concept in enumerate(conc, 1):
        row3 = idx + 2
        bg3  = "EFF6FF" if idx%2==0 else "FFFFFF"
        for col, key in enumerate(["palabra_clave","tecnologias_asociadas","metodos_marcos","areas_aplicacion"],1):
            _cell(ws3, row3, col, concept.get(key,""), bg=bg3)
        ws3.row_dimensions[row3].height = 40

    ws3.freeze_panes = "A3"

    # ── HOJA 4: SCORING ESTRATÉGICO ─────────────────────────────────
    ws4 = wb.create_sheet("Scoring Estratégico")
    ws4.merge_cells("A1:G1")
    tc4 = ws4.cell(1,1,"SCORING ESTRATÉGICO — EVALUACIÓN COMPARATIVA")
    tc4.font      = Font(bold=True, color="FFFFFF", name="Arial", size=13)
    tc4.fill      = PatternFill("solid", fgColor="0D1B2A")
    tc4.alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 32

    s4_headers = ["#","Nombre del Software","Relevancia\nAcadémica (1-5)",
                  "Valor en CV (1-5)","Uso en\nIndustria (1-5)",
                  "Diferenciación\nCompetitiva (1-5)","SCORE TOTAL"]
    s4_widths  = [4, 30, 18, 14, 16, 20, 14]
    for i,(h,w) in enumerate(zip(s4_headers, s4_widths),1):
        _hdr(ws4, 2, i, h, bg="1A3C5E")
        ws4.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws4.row_dimensions[2].height = 36

    for idx, sw in enumerate(port, 1):
        row4 = idx + 2
        bg4  = "F8FAFC" if idx%2==0 else "FFFFFF"
        capa = sw.get("capa",1)

        _cell(ws4, row4, 1, idx, halign="center", bg=bg4)
        _cell(ws4, row4, 2, sw.get("nombre",""), bold=True,
              color=CAPA_COLORS.get(capa,"000000"), bg=CAPA_BG.get(capa, bg4))

        for col, key in zip([3,4,5,6],
                            ["score_academico","score_cv","score_industria","score_diferenciacion"]):
            val = sw.get(key, 3)
            c = _cell(ws4, row4, col, val, halign="center", bg=bg4)

        # Dynamic SUM formula
        col_letters = {3:"C",4:"D",5:"E",6:"F"}
        row_ref = row4
        formula = f"=SUM(C{row_ref}:F{row_ref})"
        total_cell = ws4.cell(row4, 7, formula)
        total_cell.font      = Font(bold=True, name="Arial", size=10)
        total_cell.alignment = Alignment(horizontal="center", vertical="center")
        thin = Side(style="thin", color="E2E8F0")
        total_cell.border    = Border(left=thin,right=thin,top=thin,bottom=thin)
        total_cell.fill      = PatternFill("solid", fgColor=bg4)

        ws4.row_dimensions[row4].height = 24

    # Color scale on SCORE TOTAL column (G)
    score_range = f"G3:G{len(port)+2}"
    color_scale = ColorScaleRule(
        start_type="min", start_color="FF4444",
        mid_type="percentile", mid_value=50, mid_color="FFBB33",
        end_type="max", end_color="00C851"
    )
    ws4.conditional_formatting.add(score_range, color_scale)

    # Color scale on individual score columns C-F
    for col_letter in ["C","D","E","F"]:
        col_range = f"{col_letter}3:{col_letter}{len(port)+2}"
        ws4.conditional_formatting.add(col_range, ColorScaleRule(
            start_type="min", start_color="FFCCCC",
            end_type="max", end_color="CCFFCC"
        ))

    ws4.freeze_panes = "A3"

    # ── HOJA 5: GOOGLE DORKS ────────────────────────────────────────
    ws5 = wb.create_sheet("Google Dorks")
    ws5.merge_cells("A1:C1")
    tc5 = ws5.cell(1,1,"GOOGLE DORKS — BÚSQUEDAS AVANZADAS DE VERIFICACIÓN")
    tc5.font      = Font(bold=True, color="FFFFFF", name="Arial", size=13)
    tc5.fill      = PatternFill("solid", fgColor="166534")
    tc5.alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 32

    for i,(h,w) in enumerate(zip(["#","Google Dork","Uso Sugerido"],[6,80,40]),1):
        _hdr(ws5, 2, i, h, bg="166534")
        ws5.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws5.row_dimensions[2].height = 28

    for idx, dork in enumerate(dorks, 1):
        row5 = idx + 2
        bg5  = "F0FDF4" if idx%2==0 else "FFFFFF"
        _cell(ws5, row5, 1, idx, halign="center", bg=bg5)
        c = ws5.cell(row5, 2, dork)
        c.font      = Font(name="Courier New", size=10, color="166534")
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        thin = Side(style="thin", color="E2E8F0")
        c.border    = Border(left=thin,right=thin,top=thin,bottom=thin)
        c.fill      = PatternFill("solid", fgColor=bg5)
        _cell(ws5, row5, 3, "", bg=bg5)
        ws5.row_dimensions[row5].height = 24

    ws5.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── UI ────────────────────────────────────────────────────────────────
st.markdown("""
<div class="aria-header">
  <h1 class="aria-title">ARIA <em>Buscador</em></h1>
  <p class="aria-sub">Portafolio estratégico de software para programas académicos de posgrado</p>
</div>
""", unsafe_allow_html=True)

with st.expander("🔑 Configurar API Key de Gemini", expanded="gemini_key" not in st.session_state):
   api_key = st.secrets.get("GEMINI_API_KEY", "")

if not api_key:
    st.error("Falta configurar GEMINI_API_KEY en Streamlit Secrets.")
    st.stop()
        placeholder="AIza...",
        help="Obtén tu clave gratuita en https://aistudio.google.com/apikey"
    )
    if api_key_input:
        st.session_state["gemini_key"] = api_key_input
        st.success("✅ API Key guardada para esta sesión")

api_key = st.session_state.get("gemini_key", "")

st.markdown('<div class="section-label">01 — Dossier académico (PDF, TXT o DOCX)</div>', unsafe_allow_html=True)
uploaded = st.file_uploader(
    "Sube el dossier del programa de posgrado",
    type=["pdf","txt","docx"],
    label_visibility="collapsed"
)

st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)
st.markdown('<div class="section-label">02 — Contexto adicional (opcional)</div>', unsafe_allow_html=True)
user_context = st.text_area(
    "Contexto adicional",
    placeholder="Ej: Maestría en Administración de Hospitales, orientada a directivos de clínicas privadas en México...",
    height=90,
    label_visibility="collapsed"
)

search_btn = st.button("🔍  Generar portafolio estratégico", use_container_width=True)

if search_btn:
    if not api_key:
        st.error("⚠️ Ingresa tu API Key de Gemini para continuar.")
    elif not uploaded and not user_context:
        st.warning("Sube el dossier del programa o indica el contexto del posgrado.")
    else:
        dossier_text = extract_text(uploaded) if uploaded else ""

        with st.spinner("Analizando dossier y construyendo portafolio con Gemini... (30–60 segundos)"):
            try:
                result = search_with_gemini(api_key, dossier_text, user_context)
            except Exception as e:
                st.error(f"Error al consultar Gemini: {e}")
                st.stop()

        port = result.get("portafolio", [])
        if not port:
            st.warning("Gemini no retornó herramientas. Verifica tu dossier o contexto.")
            if result.get("_raw"):
                with st.expander("Respuesta cruda"):
                    st.text(result["_raw"])
            st.stop()

        st.session_state["result"] = result

# ── Results ───────────────────────────────────────────────────────────
if "result" in st.session_state:
    result = st.session_state["result"]
    port   = result.get("portafolio", [])
    conc   = result.get("conceptos", [])
    dorks  = result.get("google_dorks", [])

    capa1 = [s for s in port if s.get("capa")==1]
    capa2 = [s for s in port if s.get("capa")==2]
    capa3 = [s for s in port if s.get("capa")==3]

    st.markdown('<div class="aria-divider"></div>', unsafe_allow_html=True)

    col_stats, col_dl = st.columns([2,1])
    with col_stats:
        st.markdown(f"""
        <div class="stats-bar">
          <div class="stat-item"><span>{len(port)}</span> herramientas</div>
          <div class="stat-item"><span>{len(capa1)}</span> Capa 1</div>
          <div class="stat-item"><span>{len(capa2)}</span> Capa 2</div>
          <div class="stat-item"><span>{len(capa3)}</span> Capa 3</div>
          <div class="stat-item"><span>{len(conc)}</span> conceptos</div>
        </div>""", unsafe_allow_html=True)
    with col_dl:
        excel_bytes = build_excel(result)
        fname = f"ARIA_Portafolio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button("⬇️ Descargar Excel (5 hojas)", data=excel_bytes,
                           file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

    # Layer legend
    st.markdown("""
    <div class="layer-legend">
      <div class="legend-item"><div class="legend-dot" style="background:#1D4ED8"></div>Capa 1 — Core del Dossier</div>
      <div class="legend-item"><div class="legend-dot" style="background:#166534"></div>Capa 2 — Disciplinas Relacionadas</div>
      <div class="legend-item"><div class="legend-dot" style="background:#EA580C"></div>Capa 3 — Ecosistema Estratégico</div>
    </div>""", unsafe_allow_html=True)

    # Tabs for results
    tab1, tab2, tab3, tab4 = st.tabs([
        f"📦 Portafolio ({len(port)})",
        f"🗂 Paquetes",
        f"🔑 Conceptos ({len(conc)})",
        f"🔎 Google Dorks ({len(dorks)})"
    ])

    with tab1:
        for sw in port:
            capa = sw.get("capa",1)
            capa_labels = {1:"Capa 1 — Core",2:"Capa 2 — Relacionadas",3:"Capa 3 — Ecosistema"}
            badge_cls   = {1:"layer-1",2:"layer-2",3:"layer-3"}
            card_cls    = {1:"l1",2:"l2",3:"l3"}
            url = sw.get("url_oficial","")
            link_html = f'<a href="{url}" target="_blank" style="color:#1A6FA8;font-size:0.78rem;">🌐 Sitio oficial</a>' if url else ""
            score_total = (sw.get("score_academico",0)+sw.get("score_cv",0)+
                           sw.get("score_industria",0)+sw.get("score_diferenciacion",0))
            st.markdown(f"""
            <div class="sw-card {card_cls.get(capa,'l1')}">
              <div class="sw-name">{sw.get('nombre','')}</div>
              <span class="layer-badge {badge_cls.get(capa,'layer-1')}">{capa_labels.get(capa,'')}</span>
              &nbsp;<span style="font-size:0.75rem;color:#64748b;">Score: <strong>{score_total}/20</strong></span>
              <div class="sw-desc">{sw.get('valor_estrategico','')}</div>
              <div class="sw-detail"><strong>Acceso:</strong> {sw.get('tipo_acceso','')}</div>
              <div class="sw-detail"><strong>Módulo:</strong> {sw.get('relacion_dossier','')}</div>
              <div class="sw-detail"><strong>Certificación:</strong> {sw.get('certificacion','')}</div>
              <div style="margin-top:8px;">{link_html}</div>
            </div>""", unsafe_allow_html=True)

    with tab2:
        paquetes: dict = {}
        for sw in port:
            pkg = sw.get("paquete","Otros")
            paquetes.setdefault(pkg, []).append(sw)

        pkg_icons = {"Hospitales":"🏥","Medicina y Clínica":"🩺","Cuidados Médicos":"💊",
                     "Cuidados Intensivos":"🫀","Investigación y Publicación":"🔬",
                     "Dirección y Liderazgo":"📊","Salud Pública y Políticas":"🌎"}
        for pname in PAQUETES_ORDER:
            sws = paquetes.get(pname, [])
            if not sws:
                continue
            icon = pkg_icons.get(pname,"📦")
            sw_names = " · ".join(s.get("nombre","") for s in sws)
            st.markdown(f"""
            <div class="pkg-card">
              <div class="pkg-title">{icon} {pname} <span style="font-size:0.75rem;opacity:.7;">({len(sws)} herramientas)</span></div>
              <div class="pkg-obj">Equipar al estudiante con herramientas de nivel institucional en {pname.lower()}.</div>
              <div class="pkg-sw">{sw_names}</div>
              <div class="pkg-benefit">✓ Fortalece CV y empleabilidad real en el sector salud.</div>
            </div>""", unsafe_allow_html=True)

    with tab3:
        for concept in conc:
            with st.expander(f"🔑 {concept.get('palabra_clave','')}"):
                st.markdown(f"**Tecnologías asociadas:** {concept.get('tecnologias_asociadas','')}")
                st.markdown(f"**Métodos / Marcos:** {concept.get('metodos_marcos','')}")
                st.markdown(f"**Áreas de aplicación:** {concept.get('areas_aplicacion','')}")

    with tab4:
        st.caption("Copia cualquier dork y pégalo en Google para verificar herramientas")
        for i, dork in enumerate(dorks, 1):
            st.markdown(f"""
            <div class="dork-card">
              <span class="dork-num">{i}</span>
              <span class="dork-text">{dork}</span>
            </div>""", unsafe_allow_html=True)

    st.markdown(
        '<p style="text-align:center;font-size:0.72rem;color:#94a3b8;margin-top:32px;">'
        'ARIA Buscador · Portafolio Estratégico · Powered by Google Gemini 2.0 Flash</p>',
        unsafe_allow_html=True)
