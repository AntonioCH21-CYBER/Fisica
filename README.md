# ⚛️ Física IA – Chatbot para aprender física

- **Backend:** Python (Flask) + Google Gemini (gratis) + gTTS → `app.py`
- **Interfaz:** HTML / CSS / JavaScript → `web/index.html`
- **Base de datos:** SQLite en tu PC · Postgres (Neon) en Vercel
- **Base de conocimiento:** carpeta `conocimiento/` (PDF y PowerPoint del curso)

## Funciones
- 👤 Crear cuenta e iniciar sesión. Cada cuenta ve solo sus conversaciones.
- 📁 Apartados con varias conversaciones guardadas.
- 💬 Texto, 📷 imágenes, 🎤 preguntas por voz y 🔊 respuestas con audio.
- 📚 Base de conocimiento propia: responde basándose en los PDF/PPTX de `conocimiento/` y cita archivo y página.
- 📎 También se puede subir un PDF/PPTX extra a una conversación.

---

## A) Correrlo en tu PC
1. Pon tu API key en `.env` (https://aistudio.google.com/apikey).
2. Copia los PDF/PPTX del curso en `conocimiento/`.
3. En la terminal: `python app.py` → se abre http://127.0.0.1:5050

## B) Publicarlo en Vercel
Vercel no deja guardar archivos, así que las cuentas y conversaciones van en una base de datos
gratuita **Neon (Postgres)** que se conecta desde el mismo Vercel.

1. **Sube el proyecto** (GitHub o `vercel deploy`). Incluye la carpeta `conocimiento/` con tus PDF/PPTX.
   El archivo `.env` NO se sube a GitHub (está en `.gitignore`), así la clave no queda pública.
2. En tu proyecto de Vercel → pestaña **Storage** → **Create Database** → **Neon** → plan **Free** →
   **Connect** a tu proyecto. Esto agrega solo la variable `DATABASE_URL`.
3. **Settings → Environment Variables** → agrega:
   - `GEMINI_API_KEY` = tu clave de Gemini
   - (opcional) `SECRET_KEY` = cualquier texto largo y secreto
4. **Deployments** → en el último despliegue, menú **⋯ → Redeploy** (para que tome las variables).

Notas de Vercel:
- Máximo 4.4 MB por envío (fotos y audios entran sin problema; un PDF más pesado ponlo en `conocimiento/`).
- Si no conectas Neon, la web funciona pero en **modo temporal** (aparece un aviso amarillo) y los datos se borran.
