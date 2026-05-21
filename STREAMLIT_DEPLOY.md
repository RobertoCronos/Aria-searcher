# ARIA Buscador — Guía de despliegue en Streamlit Cloud

## Requisitos previos
- Cuenta en GitHub (gratuita): https://github.com
- Cuenta en Streamlit Cloud (gratuita): https://streamlit.io/cloud
- API Key de Google Gemini: https://aistudio.google.com/apikey

---

## Paso 1 — Subir el proyecto a GitHub

1. Crea un repositorio nuevo en GitHub (puede ser privado)
2. Sube todos los archivos de esta carpeta:

```bash
git init
git add .
git commit -m "ARIA Buscador inicial"
git remote add origin https://github.com/tu-usuario/aria-buscador.git
git push -u origin main
```

Estructura que debe quedar en GitHub:
```
aria-buscador/
├── app.py
├── requirements.txt
└── .streamlit/
    ├── config.toml
    └── secrets.toml     ← asegúrate de que esté en .gitignore si tiene datos
```

---

## Paso 2 — Desplegar en Streamlit Cloud

1. Ve a https://share.streamlit.io
2. Clic en **"New app"**
3. Conecta tu cuenta de GitHub si aún no lo has hecho
4. Configura:
   - **Repository:** `tu-usuario/aria-buscador`
   - **Branch:** `main`
   - **Main file path:** `app.py`
5. Clic en **"Deploy!"**

Streamlit instalará las dependencias de `requirements.txt` automáticamente.

---

## Paso 3 — Uso de la app

La API Key de Gemini se ingresa directamente en la interfaz web cada sesión.
No necesitas configurar secrets en Streamlit Cloud.

Si prefieres tenerla precargada (para uso interno):
1. En Streamlit Cloud → tu app → **Settings** → **Secrets**
2. Agrega:
```toml
GEMINI_API_KEY = "AIza..."
```
3. Y modifica en `app.py` la línea:
```python
api_key = st.session_state.get("gemini_key", "")
```
por:
```python
api_key = st.secrets.get("GEMINI_API_KEY", st.session_state.get("gemini_key", ""))
```

---

## Obtener tu API Key de Gemini (gratis)

1. Ve a https://aistudio.google.com/apikey
2. Inicia sesión con tu cuenta Google
3. Clic en **"Create API Key"**
4. Copia la clave (empieza con `AIza...`)

El plan gratuito de Gemini incluye:
- 15 requests por minuto
- 1,500 requests por día
- Suficiente para uso normal de la app

---

## Estructura de archivos

```
aria-buscador/
├── app.py                   ← App principal Streamlit + Gemini
├── requirements.txt         ← Dependencias (instaladas automáticamente)
└── .streamlit/
    ├── config.toml          ← Tema visual de la app
    └── secrets.toml         ← Solo para uso local (no subir a GitHub con datos)
```

---

## Solución de problemas

**Error "ModuleNotFoundError":**
Verifica que `requirements.txt` esté en la raíz del repositorio.

**Error de API Key inválida:**
- Verifica que empiece con `AIza`
- Asegúrate de haber habilitado la API de Gemini en Google Cloud Console

**La app no encuentra resultados:**
- Intenta con términos más específicos
- El modelo Gemini 2.0 Flash necesita internet para buscar — Streamlit Cloud lo permite por defecto
