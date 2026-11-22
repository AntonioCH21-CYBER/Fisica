# ⚛️ Física IA – Chatbot para aprender física

- **Backend:** Python (Flask) + SQLite + Google Gemini (gratis) + gTTS → `app.py`
- **Interfaz:** HTML / CSS / JavaScript → `web/index.html`
- **Tu API key:** archivo `.env`

## Funciones
- 👤 Crear cuenta e iniciar sesión (contraseñas cifradas). Cada cuenta ve solo sus conversaciones.
- 📁 Apartados (Mecánica, Laboratorio… y los que crees) con varias conversaciones guardadas.
- 💬 Chat por texto centrado en física (paso a paso, fórmulas en LaTeX), pero responde cualquier tema.
- 📷 Imágenes desde la galería o la cámara.
- 🎤 Preguntas por voz y 🔊 respuestas con audio.
- 📚 **Base de conocimiento propia**: los PDF y PowerPoint (.pptx) que pongas en la carpeta `conocimiento/`
  son la fuente principal del chatbot. En cada pregunta busca los fragmentos relevantes, responde según
  ellos y cita archivo y página/diapositiva. Debajo de cada respuesta se ve qué material consultó.
  (Los .ppt antiguos se deben guardar como .pptx o PDF. Los PDF escaneados se leen con la IA una sola vez.)
- 📎 Además, en una conversación puedes subir un PDF/PPTX extra solo para esa conversación.

## Pasos
1. Consigue tu API key gratis en https://aistudio.google.com/apikey
2. Abre `.env` y reemplaza `pega_aqui_tu_clave` por tu clave. Guarda (Ctrl + S).
3. Copia los PDF y PowerPoint del curso dentro de la carpeta `conocimiento/`.
4. Abre esta carpeta en VS Code (File → Open Folder) y en la terminal escribe:
   ```
   python app.py
   ```
   La primera vez instala solo las librerías. Luego se abre el navegador en http://127.0.0.1:5050
5. Crea tu cuenta y empieza a preguntar.

Deja la terminal abierta mientras lo usas. Para apagarlo: Ctrl + C.
Si agregas más archivos a `conocimiento/`, reinicia o usa en la web «📚 Base de conocimiento → Volver a leer la carpeta».
Todo se guarda en la carpeta `datos/` (base de datos `fisica_ia.db` y archivos subidos).
